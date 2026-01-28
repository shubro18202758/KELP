"""
Web Scraping Module with Visual Intelligence
Scrapes company websites and uses VL model for understanding visual content
"""
import asyncio
import aiohttp
import re
import json
import hashlib
import os
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
from datetime import datetime
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
import sys

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from config.settings import COMPANY_DATA_DIR


@dataclass
class WebScrapedData:
    """Data extracted from web scraping"""
    url: str
    title: str
    description: str
    products: List[str] = field(default_factory=list)
    services: List[str] = field(default_factory=list)
    achievements: List[str] = field(default_factory=list)
    metrics: List[str] = field(default_factory=list)
    certifications: List[str] = field(default_factory=list)
    news_updates: List[Dict[str, str]] = field(default_factory=list)
    leadership: List[Dict[str, str]] = field(default_factory=list)
    locations: List[str] = field(default_factory=list)
    images: List[Dict[str, str]] = field(default_factory=list)
    raw_text: str = ""
    scraped_at: str = field(default_factory=lambda: datetime.now().isoformat())
    source_urls: List[str] = field(default_factory=list)


class AsyncWebScraper:
    """
    Asynchronous web scraper optimized for company websites.
    Extracts text, images, and structured data.
    """
    
    def __init__(self, cache_dir: Path = None):
        self.cache_dir = cache_dir or (COMPANY_DATA_DIR.parent / "cache" / "web")
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.session: Optional[aiohttp.ClientSession] = None
        
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
        }
    
    async def __aenter__(self):
        self.session = aiohttp.ClientSession(headers=self.headers)
        return self
    
    async def __aexit__(self, exc_type, exc_val, exc_tb):
        if self.session:
            await self.session.close()
    
    def _get_cache_key(self, url: str) -> str:
        """Generate cache key from URL"""
        return hashlib.md5(url.encode()).hexdigest()
    
    def _get_cached(self, url: str) -> Optional[Dict]:
        """Get cached data if available and fresh"""
        cache_file = self.cache_dir / f"{self._get_cache_key(url)}.json"
        if cache_file.exists():
            try:
                with open(cache_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # Check if cache is less than 24 hours old
                cached_time = datetime.fromisoformat(data.get('scraped_at', '2000-01-01'))
                if (datetime.now() - cached_time).total_seconds() < 86400:
                    return data
            except:
                pass
        return None
    
    def _save_cache(self, url: str, data: Dict) -> None:
        """Save data to cache"""
        cache_file = self.cache_dir / f"{self._get_cache_key(url)}.json"
        try:
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
        except:
            pass
    
    async def fetch_page(self, url: str, timeout: int = 30) -> Tuple[str, int]:
        """Fetch a single page"""
        try:
            async with self.session.get(url, timeout=aiohttp.ClientTimeout(total=timeout)) as response:
                if response.status == 200:
                    html = await response.text()
                    return html, response.status
                return "", response.status
        except Exception as e:
            print(f"Error fetching {url}: {e}")
            return "", 0
    
    def _extract_text_content(self, soup: BeautifulSoup) -> str:
        """Extract meaningful text from page"""
        # Remove script, style, nav, footer elements
        for tag in soup.find_all(['script', 'style', 'nav', 'footer', 'header', 'aside']):
            tag.decompose()
        
        # Get text from main content areas
        main_content = soup.find('main') or soup.find('article') or soup.find('body')
        if main_content:
            text = main_content.get_text(separator='\n', strip=True)
            # Clean up excessive whitespace
            text = re.sub(r'\n{3,}', '\n\n', text)
            text = re.sub(r' {2,}', ' ', text)
            return text[:10000]  # Limit text length
        return ""
    
    def _extract_metadata(self, soup: BeautifulSoup) -> Dict[str, str]:
        """Extract page metadata"""
        metadata = {}
        
        # Title
        title_tag = soup.find('title')
        metadata['title'] = title_tag.get_text(strip=True) if title_tag else ""
        
        # Meta description
        meta_desc = soup.find('meta', attrs={'name': 'description'})
        metadata['description'] = meta_desc.get('content', '') if meta_desc else ""
        
        # OG tags
        og_title = soup.find('meta', property='og:title')
        if og_title:
            metadata['og_title'] = og_title.get('content', '')
        
        og_desc = soup.find('meta', property='og:description')
        if og_desc:
            metadata['og_description'] = og_desc.get('content', '')
        
        return metadata
    
    def _extract_products_services(self, soup: BeautifulSoup, text: str) -> Tuple[List[str], List[str]]:
        """Extract products and services mentions"""
        products = []
        services = []
        
        # Look for product/service sections
        product_keywords = ['product', 'solution', 'offering', 'portfolio']
        service_keywords = ['service', 'consulting', 'support', 'maintenance']
        
        # Find lists in relevant sections
        for section in soup.find_all(['section', 'div']):
            section_text = section.get_text().lower()
            
            if any(kw in section_text for kw in product_keywords):
                for li in section.find_all('li')[:10]:
                    item = li.get_text(strip=True)
                    if 5 < len(item) < 100:
                        products.append(item)
            
            if any(kw in section_text for kw in service_keywords):
                for li in section.find_all('li')[:10]:
                    item = li.get_text(strip=True)
                    if 5 < len(item) < 100:
                        services.append(item)
        
        # Extract from headings
        for heading in soup.find_all(['h2', 'h3', 'h4']):
            heading_text = heading.get_text(strip=True)
            if any(kw in heading_text.lower() for kw in product_keywords):
                products.append(heading_text)
            if any(kw in heading_text.lower() for kw in service_keywords):
                services.append(heading_text)
        
        return list(set(products))[:15], list(set(services))[:15]
    
    def _extract_metrics(self, text: str) -> List[str]:
        """Extract numerical metrics from text"""
        metrics = []
        
        # Patterns for common metrics
        patterns = [
            r'\b(\d+(?:,\d+)*(?:\.\d+)?)\s*(?:million|billion|crore|lakh)\b',
            r'\b(\d+(?:\.\d+)?)\s*%\s*(?:growth|increase|margin|CAGR)',
            r'\b(\d+)\+?\s*(?:employees?|customers?|clients?|countries?|years?)',
            r'(?:revenue|sales|profit)\s*(?:of|:)?\s*(?:₹|rs\.?|inr|usd|\$)?\s*(\d+(?:,\d+)*(?:\.\d+)?)',
            r'\b(\d{4})\s*(?:founded|established|since)',
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                context_start = max(0, match.start() - 50)
                context_end = min(len(text), match.end() + 50)
                context = text[context_start:context_end].strip()
                if len(context) < 200:
                    metrics.append(context)
        
        return list(set(metrics))[:10]
    
    def _extract_certifications(self, text: str) -> List[str]:
        """Extract certifications and awards"""
        certs = []
        
        cert_patterns = [
            r'ISO\s*\d+(?::\d+)?',
            r'IATF\s*\d+',
            r'AS\s*\d+',
            r'CMMI\s*Level\s*\d',
            r'GMP\+?',
            r'FDA\s*(?:approved|certified)?',
            r'CE\s*(?:marked|certified)?',
        ]
        
        for pattern in cert_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            certs.extend(matches)
        
        # Look for award mentions
        award_patterns = [
            r'(?:won|received|awarded)\s+(.{10,100}?award)',
            r'(best\s+.{5,50}?award)',
            r'(excellence\s+.{5,50}?award)',
        ]
        
        for pattern in award_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            certs.extend(matches)
        
        return list(set(certs))[:10]
    
    def _extract_images(self, soup: BeautifulSoup, base_url: str) -> List[Dict[str, str]]:
        """Extract relevant images from page"""
        images = []
        
        for img in soup.find_all('img'):
            src = img.get('src', '')
            alt = img.get('alt', '')
            
            if not src or 'logo' in src.lower() or 'icon' in src.lower():
                continue
            
            # Make URL absolute
            if src.startswith('//'):
                src = 'https:' + src
            elif src.startswith('/'):
                src = urljoin(base_url, src)
            elif not src.startswith('http'):
                src = urljoin(base_url, src)
            
            # Filter for quality images
            if any(ext in src.lower() for ext in ['.jpg', '.jpeg', '.png', '.webp']):
                images.append({
                    'url': src,
                    'alt': alt[:100] if alt else 'Business image',
                    'type': 'product' if 'product' in alt.lower() else 'general'
                })
        
        return images[:20]
    
    async def scrape_company(self, base_url: str, company_name: str = "") -> WebScrapedData:
        """
        Scrape a company website comprehensively.
        Visits multiple pages to gather complete information.
        """
        # Check cache first
        cached = self._get_cached(base_url)
        if cached:
            print(f"   📦 Using cached data for {base_url}")
            return WebScrapedData(**cached)
        
        print(f"   🌐 Scraping {base_url}...")
        
        result = WebScrapedData(url=base_url, title="", description="")
        all_text = []
        visited_urls = set()
        
        # Pages to visit
        pages_to_visit = [
            base_url,
            urljoin(base_url, '/about'),
            urljoin(base_url, '/about-us'),
            urljoin(base_url, '/products'),
            urljoin(base_url, '/services'),
            urljoin(base_url, '/investors'),
            urljoin(base_url, '/investor-relations'),
        ]
        
        async with aiohttp.ClientSession(headers=self.headers) as session:
            self.session = session
            
            for page_url in pages_to_visit[:5]:  # Limit to 5 pages
                if page_url in visited_urls:
                    continue
                
                visited_urls.add(page_url)
                html, status = await self.fetch_page(page_url)
                
                if status != 200 or not html:
                    continue
                
                result.source_urls.append(page_url)
                
                try:
                    soup = BeautifulSoup(html, 'lxml')
                    
                    # Extract metadata from main page
                    if page_url == base_url:
                        metadata = self._extract_metadata(soup)
                        result.title = metadata.get('title', '')
                        result.description = metadata.get('description', '')
                    
                    # Extract text
                    text = self._extract_text_content(soup)
                    all_text.append(text)
                    
                    # Extract products/services
                    products, services = self._extract_products_services(soup, text)
                    result.products.extend(products)
                    result.services.extend(services)
                    
                    # Extract metrics
                    metrics = self._extract_metrics(text)
                    result.metrics.extend(metrics)
                    
                    # Extract certifications
                    certs = self._extract_certifications(text)
                    result.certifications.extend(certs)
                    
                    # Extract images (from main page only)
                    if page_url == base_url:
                        images = self._extract_images(soup, base_url)
                        result.images = images
                    
                except Exception as e:
                    print(f"   ⚠ Error parsing {page_url}: {e}")
        
        # Combine and deduplicate
        result.raw_text = '\n\n'.join(all_text)[:20000]
        result.products = list(set(result.products))[:15]
        result.services = list(set(result.services))[:15]
        result.metrics = list(set(result.metrics))[:10]
        result.certifications = list(set(result.certifications))[:10]
        
        # Cache results
        self._save_cache(base_url, result.__dict__)
        
        return result


class IntelligentWebScraper:
    """
    Combines async scraping with VL model for intelligent extraction.
    """
    
    def __init__(self, use_vision: bool = True):
        self.async_scraper = AsyncWebScraper()
        self.use_vision = use_vision
        self._vl_engine = None
    
    def _get_vl_engine(self):
        """Lazy load VL engine"""
        if self._vl_engine is None and self.use_vision:
            try:
                from src.vision.vl_engine import get_vl_engine
                self._vl_engine = get_vl_engine()
            except Exception as e:
                print(f"VL engine not available: {e}")
                self.use_vision = False
        return self._vl_engine
    
    def scrape(self, url: str, company_name: str = "") -> WebScrapedData:
        """Synchronous wrapper for async scraper"""
        return asyncio.run(self._async_scrape(url, company_name))
    
    async def _async_scrape(self, url: str, company_name: str) -> WebScrapedData:
        """Main scraping method"""
        async with self.async_scraper as scraper:
            data = await scraper.scrape_company(url, company_name)
        
        return data
    
    def analyze_scraped_data(self, data: WebScrapedData, 
                             company_context: Dict = None) -> Dict[str, Any]:
        """
        Use LLM to extract structured insights from scraped data.
        """
        vl_engine = self._get_vl_engine()
        
        if vl_engine:
            prompt = f"""Analyze this company information and extract key business insights:

Company Title: {data.title}
Description: {data.description}

Raw Content (excerpt):
{data.raw_text[:3000]}

Products Found: {', '.join(data.products[:10])}
Services Found: {', '.join(data.services[:10])}
Metrics Found: {data.metrics}
Certifications: {data.certifications}

Extract and structure:
1. Core business model (1 sentence)
2. Key competitive advantages (3-5 points)
3. Market positioning
4. Growth indicators
5. Any unique selling propositions

Be specific and use the actual data provided."""

            response = vl_engine.generate_creative_content(prompt)
            
            return {
                "structured_insights": response,
                "products": data.products,
                "services": data.services,
                "metrics": data.metrics,
                "certifications": data.certifications,
                "images": data.images,
                "source_urls": data.source_urls
            }
        
        # Fallback without VL
        return {
            "products": data.products,
            "services": data.services,
            "metrics": data.metrics,
            "certifications": data.certifications,
            "images": data.images,
            "source_urls": data.source_urls
        }


def scrape_company_website(url: str, company_name: str = "") -> WebScrapedData:
    """
    Main function to scrape a company website.
    """
    if not url or url.lower() == 'not available':
        return WebScrapedData(url="", title="", description="")
    
    # Ensure URL has scheme
    if not url.startswith('http'):
        url = 'https://' + url
    
    scraper = IntelligentWebScraper(use_vision=False)  # Start without vision for speed
    return scraper.scrape(url, company_name)


if __name__ == "__main__":
    print("Testing Web Scraper...\n")
    
    # Test with a sample URL
    test_urls = [
        "https://www.kalyaniforge.com",
        "https://www.centumelectronics.com",
    ]
    
    for url in test_urls[:1]:  # Test with first URL
        print(f"Scraping: {url}")
        data = scrape_company_website(url)
        
        print(f"\nResults:")
        print(f"  Title: {data.title}")
        print(f"  Description: {data.description[:100]}..." if data.description else "  No description")
        print(f"  Products: {len(data.products)}")
        print(f"  Services: {len(data.services)}")
        print(f"  Metrics: {len(data.metrics)}")
        print(f"  Certifications: {data.certifications}")
        print(f"  Images: {len(data.images)}")
        print(f"  Pages scraped: {len(data.source_urls)}")
