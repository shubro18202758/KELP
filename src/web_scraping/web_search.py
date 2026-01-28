"""
Intelligent Web Search & Scraping Module
=========================================
Real web search using DuckDuckGo and intelligent content extraction.
Used for:
- Public data layer: Business model, products, market sentiment
- Company website crawling
- News and blog extraction
- Visual intelligence: Finding relevant images
"""
import os
import re
import json
import asyncio
import aiohttp
import hashlib
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
from datetime import datetime
from urllib.parse import urljoin, urlparse, quote_plus
from bs4 import BeautifulSoup, Tag
import sys

sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from config.settings import OUTPUT_DIR


@dataclass
class SearchResult:
    """Represents a single search result"""
    title: str
    url: str
    snippet: str
    source: str = ""
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())


@dataclass
class ExtractedContent:
    """Extracted content from a webpage"""
    url: str
    title: str
    main_content: str
    headings: List[str]
    key_facts: List[str]
    images: List[Dict[str, str]]
    metadata: Dict[str, Any]
    extraction_time: str = field(default_factory=lambda: datetime.now().isoformat())


@dataclass  
class CompanyWebData:
    """Aggregated web data for a company"""
    company_name: str
    search_results: List[SearchResult]
    extracted_pages: List[ExtractedContent]
    images: List[Dict[str, str]]
    news_items: List[Dict[str, str]]
    key_insights: List[str]
    sources_used: List[str]


class DuckDuckGoSearch:
    """
    DuckDuckGo search interface - no API key required.
    Uses the HTML search to extract results with retry logic.
    """
    
    BASE_URL = "https://html.duckduckgo.com/html/"
    LITE_URL = "https://lite.duckduckgo.com/lite/"
    
    # Multiple user agents for rotation
    USER_AGENTS = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Safari/605.1.15',
        'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    ]
    
    def __init__(self):
        self.session = None
        self._ua_index = 0
    
    def _get_headers(self) -> Dict:
        """Get headers with rotating user agent"""
        ua = self.USER_AGENTS[self._ua_index % len(self.USER_AGENTS)]
        self._ua_index += 1
        return {
            'User-Agent': ua,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
        }
    
    async def _get_session(self) -> aiohttp.ClientSession:
        if self.session is None or self.session.closed:
            timeout = aiohttp.ClientTimeout(total=30)
            connector = aiohttp.TCPConnector(limit=5, ssl=False)
            self.session = aiohttp.ClientSession(timeout=timeout, connector=connector)
        return self.session
    
    async def search(self, query: str, max_results: int = 10, retries: int = 3) -> List[SearchResult]:
        """
        Perform a DuckDuckGo search with retry logic.
        
        Args:
            query: Search query
            max_results: Maximum results to return
            retries: Number of retry attempts
            
        Returns:
            List of SearchResult objects
        """
        session = await self._get_session()
        results = []
        
        for attempt in range(retries):
            try:
                # Prepare search request with rotating headers
                data = {'q': query, 'b': ''}
                headers = self._get_headers()
                
                # Try HTML endpoint first, then lite
                url = self.BASE_URL if attempt < 2 else self.LITE_URL
                
                async with session.post(url, data=data, headers=headers) as response:
                    if response.status == 202:
                        # Rate limited, wait and retry
                        await asyncio.sleep(2 * (attempt + 1))
                        continue
                    elif response.status != 200:
                        continue
                    
                    html = await response.text()
                    results = self._parse_results(html, max_results)
                    if results:
                        return results
                
            except Exception as e:
                if attempt < retries - 1:
                    await asyncio.sleep(1)
                    continue
        
        # If DDG fails, try direct company website search
        return await self._fallback_search(query, max_results)
    
    def _parse_results(self, html: str, max_results: int) -> List[SearchResult]:
        """Parse HTML search results"""
        results = []
        soup = BeautifulSoup(html, 'lxml')
        
        for result in soup.select('.result, .result-link'):
            if len(results) >= max_results:
                break
            
            # Extract title and URL
            title_elem = result.select_one('.result__title a, a.result-link')
            if not title_elem:
                continue
            
            title = title_elem.get_text(strip=True)
            url = title_elem.get('href', '')
            
            # Extract snippet
            snippet_elem = result.select_one('.result__snippet, .result-snippet')
            snippet = snippet_elem.get_text(strip=True) if snippet_elem else ""
            
            # Extract source
            source_elem = result.select_one('.result__url')
            source = source_elem.get_text(strip=True) if source_elem else urlparse(url).netloc
            
            if url and title:
                results.append(SearchResult(
                    title=title,
                    url=url,
                    snippet=snippet,
                    source=source
                ))
        
        return results
    
    async def _fallback_search(self, query: str, max_results: int) -> List[SearchResult]:
        """Fallback: Try direct company page search"""
        results = []
        
        # Extract company name and create potential URLs
        words = query.replace('"', '').split()
        company_name = '_'.join(words[:2]).lower()
        
        potential_urls = [
            f"https://www.{company_name.replace('_', '')}.com",
            f"https://www.linkedin.com/company/{company_name.replace('_', '-')}",
            f"https://craft.co/{company_name.replace('_', '-')}",
        ]
        
        for url in potential_urls[:max_results]:
            results.append(SearchResult(
                title=f"Potential: {url}",
                url=url,
                snippet="Direct URL attempt",
                source=urlparse(url).netloc
            ))
        
        return results
    
    async def search_with_cache(self, query: str, cache: Dict = None, max_results: int = 10) -> List[SearchResult]:
        """Search with caching to avoid repeated queries"""
        if cache is not None and query in cache:
            return cache[query]
        
        results = await self.search(query, max_results)
        
        if cache is not None:
            cache[query] = results
        
        return results
    
    async def search_images(self, query: str, max_results: int = 5) -> List[Dict[str, str]]:
        """
        Search for images using DuckDuckGo.
        
        Args:
            query: Image search query
            max_results: Maximum images to return
            
        Returns:
            List of image info dicts with url, title, source
        """
        session = await self._get_session()
        images = []
        
        try:
            # Use image search endpoint
            search_url = f"https://duckduckgo.com/?q={quote_plus(query)}&iax=images&ia=images"
            
            # For now, return placeholder - full image search requires JS
            # We'll use the web scraper to find images on company pages instead
            return images
            
        except Exception as e:
            print(f"   ⚠ Image search error: {e}")
            return images
    
    async def close(self):
        if self.session and not self.session.closed:
            await self.session.close()


class IntelligentScraper:
    """
    Intelligent web scraper with content extraction.
    Focuses on extracting business-relevant information.
    """
    
    def __init__(self):
        self.session = None
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        }
        self.cache_dir = OUTPUT_DIR / "web_cache"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        
    async def _get_session(self) -> aiohttp.ClientSession:
        if self.session is None or self.session.closed:
            timeout = aiohttp.ClientTimeout(total=30)
            connector = aiohttp.TCPConnector(limit=10, ssl=False)
            self.session = aiohttp.ClientSession(
                timeout=timeout,
                connector=connector,
                headers=self.headers
            )
        return self.session
    
    def _get_cache_path(self, url: str) -> Path:
        """Get cache file path for URL"""
        url_hash = hashlib.md5(url.encode()).hexdigest()
        return self.cache_dir / f"{url_hash}.json"
    
    def _load_from_cache(self, url: str) -> Optional[ExtractedContent]:
        """Load cached content if available"""
        cache_path = self._get_cache_path(url)
        if cache_path.exists():
            try:
                with open(cache_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return ExtractedContent(**data)
            except:
                pass
        return None
    
    def _save_to_cache(self, content: ExtractedContent):
        """Save content to cache"""
        cache_path = self._get_cache_path(content.url)
        try:
            with open(cache_path, 'w', encoding='utf-8') as f:
                json.dump({
                    'url': content.url,
                    'title': content.title,
                    'main_content': content.main_content,
                    'headings': content.headings,
                    'key_facts': content.key_facts,
                    'images': content.images,
                    'metadata': content.metadata,
                    'extraction_time': content.extraction_time
                }, f)
        except:
            pass
    
    async def extract_page(self, url: str, use_cache: bool = True) -> Optional[ExtractedContent]:
        """
        Extract content from a webpage.
        
        Args:
            url: URL to scrape
            use_cache: Whether to use cached content
            
        Returns:
            ExtractedContent or None
        """
        # Check cache first
        if use_cache:
            cached = self._load_from_cache(url)
            if cached:
                return cached
        
        session = await self._get_session()
        
        try:
            async with session.get(url) as response:
                if response.status != 200:
                    return None
                html = await response.text()
            
            soup = BeautifulSoup(html, 'lxml')
            
            # Extract title
            title = ""
            title_tag = soup.find('title')
            if title_tag:
                title = title_tag.get_text(strip=True)
            
            # Remove script and style elements
            for element in soup(['script', 'style', 'nav', 'footer', 'header', 'aside']):
                element.decompose()
            
            # Extract headings
            headings = []
            for h in soup.find_all(['h1', 'h2', 'h3']):
                text = h.get_text(strip=True)
                if text and len(text) > 3:
                    headings.append(text)
            
            # Extract main content
            main_content = self._extract_main_content(soup)
            
            # Extract key facts (numbers, dates, etc.)
            key_facts = self._extract_key_facts(main_content)
            
            # Extract images
            images = self._extract_images(soup, url)
            
            # Extract metadata
            metadata = self._extract_metadata(soup)
            
            content = ExtractedContent(
                url=url,
                title=title,
                main_content=main_content,
                headings=headings,
                key_facts=key_facts,
                images=images,
                metadata=metadata
            )
            
            # Cache it
            self._save_to_cache(content)
            
            return content
            
        except Exception as e:
            print(f"   ⚠ Scraping error for {url}: {e}")
            return None
    
    def _extract_main_content(self, soup: BeautifulSoup) -> str:
        """Extract main text content from page"""
        # Try to find main content area
        main_areas = ['main', 'article', '.content', '#content', '.main-content', '[role="main"]']
        
        main_elem = None
        for selector in main_areas:
            main_elem = soup.select_one(selector)
            if main_elem:
                break
        
        if not main_elem:
            main_elem = soup.find('body')
        
        if not main_elem:
            return ""
        
        # Get all text
        paragraphs = []
        for p in main_elem.find_all(['p', 'li', 'div']):
            text = p.get_text(strip=True)
            if len(text) > 50:  # Filter short snippets
                paragraphs.append(text)
        
        return "\n\n".join(paragraphs[:30])  # Limit to first 30 paragraphs
    
    def _extract_key_facts(self, content: str) -> List[str]:
        """Extract key facts from content"""
        facts = []
        
        # Pattern for numbers with context
        patterns = [
            r'\d+(?:\.\d+)?%[^.]*',  # Percentages
            r'(?:₹|Rs\.?|USD|\$)\s*[\d,]+(?:\.\d+)?[^.]*',  # Currency
            r'\d{4}\s*-\s*\d{4}',  # Year ranges
            r'(?:founded|established|since)\s+\d{4}',  # Founding years
            r'\d+(?:,\d{3})*\+?\s+(?:employees?|workers?|staff|people)',  # Employee counts
            r'\d+(?:,\d{3})*\+?\s+(?:clients?|customers?|users?)',  # Customer counts
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, content, re.IGNORECASE)
            facts.extend(matches[:3])  # Limit per pattern
        
        return list(set(facts))[:15]  # Dedupe and limit
    
    def _extract_images(self, soup: BeautifulSoup, base_url: str) -> List[Dict[str, str]]:
        """Extract relevant images from page"""
        images = []
        
        for img in soup.find_all('img'):
            src = img.get('src') or img.get('data-src')
            if not src:
                continue
            
            # Make absolute URL
            if not src.startswith(('http://', 'https://')):
                src = urljoin(base_url, src)
            
            # Filter out small images and icons
            width = img.get('width', '0')
            height = img.get('height', '0')
            try:
                if int(width) < 100 or int(height) < 100:
                    continue
            except:
                pass
            
            # Get alt text
            alt = img.get('alt', '')
            
            # Skip common non-content images
            skip_patterns = ['logo', 'icon', 'avatar', 'social', 'button', 'arrow']
            if any(pattern in src.lower() or pattern in alt.lower() for pattern in skip_patterns):
                continue
            
            images.append({
                'url': src,
                'alt': alt,
                'title': img.get('title', '')
            })
        
        return images[:10]  # Limit to 10 images
    
    def _extract_metadata(self, soup: BeautifulSoup) -> Dict[str, Any]:
        """Extract page metadata"""
        metadata = {}
        
        # Meta description
        desc = soup.find('meta', attrs={'name': 'description'})
        if desc:
            metadata['description'] = desc.get('content', '')
        
        # Keywords
        keywords = soup.find('meta', attrs={'name': 'keywords'})
        if keywords:
            metadata['keywords'] = keywords.get('content', '')
        
        # Open Graph
        for og in soup.find_all('meta', attrs={'property': re.compile(r'^og:')}):
            prop = og.get('property', '').replace('og:', '')
            metadata[f'og_{prop}'] = og.get('content', '')
        
        return metadata
    
    async def close(self):
        if self.session and not self.session.closed:
            await self.session.close()


class WebSearchPipeline:
    """
    Complete web search and scraping pipeline for company research.
    """
    
    def __init__(self):
        self.search = DuckDuckGoSearch()
        self.scraper = IntelligentScraper()
        self.output_dir = OUTPUT_DIR / "web_data"
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    async def research_company(self, company_name: str, sector: str = "") -> CompanyWebData:
        """
        Conduct comprehensive web research on a company.
        
        Args:
            company_name: Name of the company
            sector: Industry sector (optional, helps with search)
            
        Returns:
            CompanyWebData with all gathered information
        """
        print(f"   🔍 Researching: {company_name}")
        
        # Build search queries
        queries = [
            f'"{company_name}" company profile',
            f'"{company_name}" products services',
            f'"{company_name}" {sector} business' if sector else f'"{company_name}" business',
            f'"{company_name}" financials revenue',
            f'"{company_name}" news recent'
        ]
        
        # Perform searches
        all_results = []
        for query in queries[:3]:  # Limit queries
            print(f"      Searching: {query[:40]}...")
            results = await self.search.search(query, max_results=5)
            all_results.extend(results)
            await asyncio.sleep(1)  # Rate limiting
        
        # Dedupe by URL
        seen_urls = set()
        unique_results = []
        for r in all_results:
            if r.url not in seen_urls:
                seen_urls.add(r.url)
                unique_results.append(r)
        
        print(f"      Found {len(unique_results)} unique results")
        
        # Extract content from top pages
        extracted_pages = []
        for result in unique_results[:5]:  # Limit pages to scrape
            print(f"      Extracting: {result.source}...")
            content = await self.scraper.extract_page(result.url)
            if content:
                extracted_pages.append(content)
            await asyncio.sleep(0.5)
        
        # Collect images from all pages
        all_images = []
        for page in extracted_pages:
            all_images.extend(page.images)
        
        # Dedupe images
        seen_image_urls = set()
        unique_images = []
        for img in all_images:
            if img['url'] not in seen_image_urls:
                seen_image_urls.add(img['url'])
                unique_images.append(img)
        
        # Extract news items (look for news-related content)
        news_items = []
        news_keywords = ['news', 'press', 'announcement', 'update', 'launch', 'partnership']
        for page in extracted_pages:
            for heading in page.headings:
                if any(kw in heading.lower() for kw in news_keywords):
                    news_items.append({
                        'title': heading,
                        'source': page.url
                    })
        
        # Generate key insights
        key_insights = self._generate_insights(extracted_pages)
        
        # Collect sources
        sources = [page.url for page in extracted_pages]
        
        web_data = CompanyWebData(
            company_name=company_name,
            search_results=unique_results,
            extracted_pages=extracted_pages,
            images=unique_images[:10],
            news_items=news_items[:5],
            key_insights=key_insights,
            sources_used=sources
        )
        
        # Save to file
        self._save_web_data(web_data)
        
        print(f"      ✓ Research complete: {len(extracted_pages)} pages, {len(unique_images)} images")
        
        return web_data
    
    def _generate_insights(self, pages: List[ExtractedContent]) -> List[str]:
        """Generate key insights from extracted pages"""
        insights = []
        
        # Collect all key facts
        all_facts = []
        for page in pages:
            all_facts.extend(page.key_facts)
        
        # Dedupe and clean
        seen = set()
        for fact in all_facts:
            clean_fact = fact.strip()
            if clean_fact and clean_fact not in seen and len(clean_fact) > 10:
                seen.add(clean_fact)
                insights.append(clean_fact)
        
        return insights[:10]
    
    def _save_web_data(self, data: CompanyWebData):
        """Save web data to JSON file"""
        filename = f"{data.company_name.lower().replace(' ', '_')}_web_data.json"
        output_path = self.output_dir / filename
        
        # Convert to serializable format
        data_dict = {
            'company_name': data.company_name,
            'search_results': [
                {'title': r.title, 'url': r.url, 'snippet': r.snippet, 'source': r.source}
                for r in data.search_results
            ],
            'extracted_pages': [
                {
                    'url': p.url,
                    'title': p.title,
                    'content_preview': p.main_content[:500] + '...' if len(p.main_content) > 500 else p.main_content,
                    'headings': p.headings,
                    'key_facts': p.key_facts
                }
                for p in data.extracted_pages
            ],
            'images': data.images,
            'news_items': data.news_items,
            'key_insights': data.key_insights,
            'sources_used': data.sources_used
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data_dict, f, indent=2, ensure_ascii=False)
    
    async def search_sector_images(self, sector: str, 
                                   keywords: List[str] = None) -> List[Dict[str, str]]:
        """
        Search for generic sector images (no company branding).
        
        Args:
            sector: Industry sector
            keywords: Additional keywords
            
        Returns:
            List of image info dicts
        """
        queries = [
            f"{sector} industry professional photo",
            f"{sector} manufacturing facility generic",
            f"{sector} products professional",
        ]
        
        if keywords:
            for kw in keywords[:2]:
                queries.append(f"{sector} {kw} professional photo")
        
        images = []
        for query in queries:
            results = await self.search.search_images(query, max_results=3)
            images.extend(results)
        
        return images
    
    async def close(self):
        await self.search.close()
        await self.scraper.close()


async def research_company_async(company_name: str, sector: str = "") -> CompanyWebData:
    """
    Convenience function for researching a company.
    """
    pipeline = WebSearchPipeline()
    try:
        return await pipeline.research_company(company_name, sector)
    finally:
        await pipeline.close()


def research_company(company_name: str, sector: str = "") -> CompanyWebData:
    """
    Synchronous wrapper for company research.
    """
    return asyncio.run(research_company_async(company_name, sector))


if __name__ == "__main__":
    print("Testing Web Search Pipeline...")
    print("=" * 50)
    
    async def test():
        pipeline = WebSearchPipeline()
        
        try:
            # Test search
            print("\n1. Testing DuckDuckGo Search...")
            results = await pipeline.search.search("Kalyani Forge automotive", max_results=5)
            print(f"   Found {len(results)} results")
            for r in results[:3]:
                print(f"   - {r.title[:50]}...")
            
            # Test page extraction
            print("\n2. Testing Page Extraction...")
            if results:
                content = await pipeline.scraper.extract_page(results[0].url)
                if content:
                    print(f"   Title: {content.title[:50]}...")
                    print(f"   Headings: {len(content.headings)}")
                    print(f"   Key facts: {len(content.key_facts)}")
                    print(f"   Images: {len(content.images)}")
            
            # Test full research
            print("\n3. Testing Full Research...")
            web_data = await pipeline.research_company("Ksolves India", "Technology")
            print(f"   Pages extracted: {len(web_data.extracted_pages)}")
            print(f"   Images found: {len(web_data.images)}")
            print(f"   Key insights: {len(web_data.key_insights)}")
            
        finally:
            await pipeline.close()
    
    asyncio.run(test())
