"""
Advanced Research Engine - Gemini-Style Web Research
=====================================================

Implements deep web research with:
1. Multi-source search (DuckDuckGo via ddgs package)
2. Actual webpage content fetching and extraction
3. LLM-powered summarization with source citations
4. Statistics extraction and validation
5. Market intelligence aggregation

This replicates Gemini's ability to search, read, and synthesize web content.
"""

import re
import json
import asyncio
import aiohttp
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
from urllib.parse import quote_plus, urlparse
from bs4 import BeautifulSoup
import time

# Import new ddgs package for DuckDuckGo search
try:
    from ddgs import DDGS
    HAS_DDGS = True
except ImportError:
    HAS_DDGS = False
    print("⚠ ddgs package not installed. Run: pip install ddgs")


@dataclass
class WebSource:
    """A single web source with extracted content"""
    url: str
    title: str
    domain: str
    content: str  # Full extracted text
    snippet: str  # Brief summary
    statistics: List[str] = field(default_factory=list)
    relevance_score: float = 0.0
    fetch_time: float = 0.0


@dataclass
class MarketIntelligence:
    """Comprehensive market intelligence from web research"""
    # Market sizing
    market_size: str = ""
    market_size_year: str = ""
    market_cagr: str = ""
    tam_sam_som: Dict[str, str] = field(default_factory=dict)
    
    # Competitive landscape
    key_players: List[str] = field(default_factory=list)
    market_shares: Dict[str, str] = field(default_factory=dict)
    
    # Industry trends
    trends: List[str] = field(default_factory=list)
    growth_drivers: List[str] = field(default_factory=list)
    challenges: List[str] = field(default_factory=list)
    
    # Financial benchmarks
    industry_margins: Dict[str, str] = field(default_factory=dict)
    valuation_multiples: Dict[str, str] = field(default_factory=dict)
    
    # Raw statistics extracted
    statistics: Dict[str, str] = field(default_factory=dict)
    
    # Sources for citation
    sources: List[str] = field(default_factory=list)
    
    # LLM-generated insights
    executive_summary: str = ""
    investment_implications: List[str] = field(default_factory=list)


class AdvancedResearchEngine:
    """
    Gemini-style web research engine.
    
    Key capabilities:
    1. Search multiple sources with smart query generation
    2. Fetch and parse actual webpage content
    3. Extract statistics, facts, and figures
    4. LLM-powered synthesis and summarization
    5. Source attribution and citation
    """
    
    def __init__(self, ollama_url: str = "http://localhost:11434"):
        self.ollama_url = ollama_url
        self.model = "qwen2.5:7b"
        self.session: Optional[aiohttp.ClientSession] = None
        
        # Request headers to avoid blocking
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
        }
        
        # Rate limiting
        self._last_request_time = 0
        self._min_request_interval = 0.5  # seconds
        
    async def _get_session(self) -> aiohttp.ClientSession:
        """Get or create HTTP session"""
        if self.session is None or self.session.closed:
            timeout = aiohttp.ClientTimeout(total=30, connect=10)
            connector = aiohttp.TCPConnector(limit=10, limit_per_host=3)
            self.session = aiohttp.ClientSession(
                timeout=timeout, 
                connector=connector,
                headers=self.headers
            )
        return self.session
    
    async def close(self):
        """Close the HTTP session"""
        if self.session and not self.session.closed:
            await self.session.close()
    
    async def _rate_limit(self):
        """Simple rate limiting"""
        now = time.time()
        elapsed = now - self._last_request_time
        if elapsed < self._min_request_interval:
            await asyncio.sleep(self._min_request_interval - elapsed)
        self._last_request_time = time.time()
    
    # =========================================================================
    # SEARCH FUNCTIONALITY
    # =========================================================================
    
    async def search_duckduckgo(self, query: str, num_results: int = 10) -> List[Dict]:
        """Search DuckDuckGo using the ddgs package"""
        await self._rate_limit()
        results = []
        
        if not HAS_DDGS:
            print("  ⚠ ddgs package not available")
            return results
        
        try:
            # Use the ddgs package - runs synchronously but wrapped in async
            ddgs_results = list(DDGS().text(query, max_results=num_results))
            
            for r in ddgs_results:
                url = r.get('href', '')
                title = r.get('title', '')
                snippet = r.get('body', '')
                
                if url and title:
                    results.append({
                        'url': url,
                        'title': title,
                        'snippet': snippet,
                        'domain': urlparse(url).netloc
                    })
                    
        except Exception as e:
            print(f"  ⚠ Search error: {e}")
        
        return results
    
    async def fetch_webpage_content(self, url: str, max_chars: int = 15000) -> Optional[WebSource]:
        """
        Fetch and extract readable content from a webpage.
        This is key - we actually READ the pages like Gemini does.
        """
        await self._rate_limit()
        start_time = time.time()
        
        try:
            session = await self._get_session()
            
            async with session.get(url, allow_redirects=True) as resp:
                if resp.status != 200:
                    return None
                
                content_type = resp.headers.get('content-type', '')
                if 'text/html' not in content_type and 'text/plain' not in content_type:
                    return None
                
                html = await resp.text()
                soup = BeautifulSoup(html, 'html.parser')
                
                # Remove script, style, nav, footer, header elements
                for tag in soup(['script', 'style', 'nav', 'footer', 'header', 
                                'aside', 'iframe', 'noscript', 'form']):
                    tag.decompose()
                
                # Get title
                title = ""
                title_tag = soup.find('title')
                if title_tag:
                    title = title_tag.get_text(strip=True)
                
                # Extract main content
                # Try common content containers first
                main_content = None
                for selector in ['article', 'main', '.content', '.post-content', 
                                '.article-body', '#content', '.entry-content']:
                    main_content = soup.select_one(selector)
                    if main_content:
                        break
                
                if not main_content:
                    main_content = soup.find('body')
                
                if not main_content:
                    return None
                
                # Get text content
                text = main_content.get_text(separator=' ', strip=True)
                text = re.sub(r'\s+', ' ', text)  # Normalize whitespace
                text = text[:max_chars]
                
                # Extract statistics from the content
                statistics = self._extract_statistics(text)
                
                fetch_time = time.time() - start_time
                
                return WebSource(
                    url=url,
                    title=title,
                    domain=urlparse(url).netloc,
                    content=text,
                    snippet=text[:300] + "..." if len(text) > 300 else text,
                    statistics=statistics,
                    fetch_time=fetch_time
                )
                
        except asyncio.TimeoutError:
            print(f"  ⚠ Timeout fetching: {urlparse(url).netloc}")
        except Exception as e:
            pass  # Silently skip failed pages
        
        return None
    
    def _extract_statistics(self, text: str) -> List[str]:
        """Extract meaningful statistics from text"""
        stats = []
        
        patterns = [
            # Market size patterns
            (r'(?:market\s+size|market\s+valued|worth|estimated)\s*(?:at|of)?\s*'
             r'[\$₹]?\s*[\d,.]+\s*(?:billion|million|trillion|crore|Bn|B|M|Cr)', 'market_size'),
            
            # CAGR patterns
            (r'CAGR\s*(?:of)?\s*[\d.]+\s*%', 'cagr'),
            (r'(?:growing|growth)\s*(?:at|of)?\s*[\d.]+\s*%\s*(?:CAGR|annually)?', 'growth'),
            
            # Revenue/financial
            (r'(?:revenue|sales|turnover)\s*(?:of|at|reached)?\s*[\$₹]?\s*[\d,.]+\s*'
             r'(?:billion|million|crore|Cr|B|M)', 'revenue'),
            
            # Margin patterns
            (r'(?:EBITDA|operating|profit|net)\s*margin\s*(?:of|at)?\s*[\d.]+\s*%', 'margin'),
            
            # Market share
            (r'(?:market\s+share|share)\s*(?:of|at)?\s*[\d.]+\s*%', 'share'),
            
            # Year projections
            (r'(?:by|in|reach)\s*(?:20\d{2})\s*[\$₹]?\s*[\d,.]+\s*'
             r'(?:billion|million|trillion|crore)', 'projection'),
            
            # Employee/scale
            (r'[\d,]+\s*(?:employees|workforce|staff|professionals)', 'scale'),
            
            # Percentage changes
            (r'(?:increased|grew|rose|declined)\s*(?:by)?\s*[\d.]+\s*%', 'change'),
        ]
        
        for pattern, stat_type in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches[:3]:  # Limit per type
                stats.append(match.strip())
        
        return list(set(stats))[:15]  # Dedupe and limit
    
    # =========================================================================
    # LLM INTEGRATION FOR ANALYSIS
    # =========================================================================
    
    async def _call_llm(self, prompt: str, max_tokens: int = 1500, 
                        temperature: float = 0.3) -> str:
        """Call local LLM with optimized parameters for factual extraction"""
        try:
            async with aiohttp.ClientSession(
                timeout=aiohttp.ClientTimeout(total=90)
            ) as session:
                payload = {
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": temperature,  # Low for factual content
                        "num_predict": max_tokens,
                        "num_gpu": 99,
                        "top_p": 0.9,
                        "repeat_penalty": 1.15,
                        "num_ctx": 4096,  # Larger context
                    }
                }
                
                async with session.post(
                    f"{self.ollama_url}/api/generate", 
                    json=payload
                ) as resp:
                    if resp.status == 200:
                        data = await resp.json()
                        return data.get("response", "").strip()
        except Exception as e:
            print(f"  ⚠ LLM call error: {e}")
        
        return ""
    
    async def synthesize_research(self, sources: List[WebSource], 
                                   sector: str) -> Dict[str, Any]:
        """
        Gemini-style synthesis: Read all sources and generate insights.
        """
        if not sources:
            return {}
        
        # Combine content from all sources
        combined_content = "\n\n---\n\n".join([
            f"SOURCE: {s.domain}\nTITLE: {s.title}\nCONTENT:\n{s.content[:3000]}"
            for s in sources[:5]  # Top 5 sources
        ])
        
        prompt = f"""You are a senior M&A research analyst synthesizing web research for an investment teaser.

SECTOR: {sector}

WEB RESEARCH SOURCES:
{combined_content}

TASK: Extract and synthesize KEY INVESTMENT-RELEVANT DATA from these sources.

OUTPUT FORMAT (JSON):
{{
    "market_size": "exact figure with currency and year, e.g., '$45 Billion (2024)'",
    "market_cagr": "percentage with period, e.g., '12.5% CAGR (2024-2030)'",
    "key_trends": ["trend 1 with specific data", "trend 2 with numbers", "trend 3"],
    "growth_drivers": ["driver 1 with quantification", "driver 2", "driver 3"],
    "key_statistics": {{
        "stat_name": "stat_value with source"
    }},
    "competitive_landscape": "brief overview with market share data if available",
    "investment_implications": ["implication 1 for investors", "implication 2"]
}}

RULES:
1. ONLY include data that appears in the sources - do not fabricate
2. Include SPECIFIC NUMBERS wherever available
3. Cite approximate source (e.g., "per industry reports")
4. If data is not available, use null instead of making up values
5. Focus on India market data when available

OUTPUT JSON:"""

        response = await self._call_llm(prompt, max_tokens=1200, temperature=0.2)
        
        try:
            # Extract JSON from response
            json_match = re.search(r'\{[\s\S]+\}', response)
            if json_match:
                return json.loads(json_match.group())
        except json.JSONDecodeError:
            pass
        
        return {}
    
    # =========================================================================
    # COMPREHENSIVE RESEARCH PIPELINE
    # =========================================================================
    
    async def deep_research(self, sector: str, sub_sector: str = "",
                            company_context: str = "") -> MarketIntelligence:
        """
        Perform comprehensive Gemini-style research.
        
        1. Generate smart search queries
        2. Search and rank results
        3. Fetch actual page content
        4. Extract statistics
        5. LLM synthesis
        6. Return structured intelligence
        """
        intel = MarketIntelligence()
        
        print(f"  🔍 Deep researching: {sector}")
        
        # Generate targeted search queries
        queries = [
            f"{sector} India market size 2024 2025 billion",
            f"{sector} industry CAGR growth rate forecast",
            f"{sector} {sub_sector} market trends analysis",
            f"{sector} India top companies market share",
            f"{sector} industry outlook investment opportunities",
        ]
        
        # If we have company context, add more specific queries
        if sub_sector:
            queries.append(f"{sub_sector} market size India growth")
        
        # Collect all search results
        all_results = []
        for query in queries[:4]:  # Limit to 4 queries
            results = await self.search_duckduckgo(query, 6)
            all_results.extend(results)
        
        # Deduplicate by URL
        seen_urls = set()
        unique_results = []
        for r in all_results:
            if r['url'] not in seen_urls:
                seen_urls.add(r['url'])
                unique_results.append(r)
        
        print(f"  📄 Found {len(unique_results)} unique sources")
        
        # Fetch actual page content (parallel with limit)
        fetch_tasks = [
            self.fetch_webpage_content(r['url']) 
            for r in unique_results[:8]  # Limit to 8 pages
        ]
        
        fetched_sources = await asyncio.gather(*fetch_tasks, return_exceptions=True)
        
        # Filter successful fetches
        valid_sources = [
            s for s in fetched_sources 
            if isinstance(s, WebSource) and s.content
        ]
        
        print(f"  📖 Successfully read {len(valid_sources)} pages")
        
        # Extract statistics from snippets (fallback)
        for result in unique_results:
            stats = self._extract_statistics(result['snippet'])
            for stat in stats:
                if 'market' in stat.lower() and not intel.market_size:
                    intel.market_size = stat
                elif 'cagr' in stat.lower() and not intel.market_cagr:
                    intel.market_cagr = stat
        
        # Extract from fetched content
        for source in valid_sources:
            intel.sources.append(f"{source.domain}: {source.title[:50]}")
            
            for stat in source.statistics:
                stat_lower = stat.lower()
                if 'market' in stat_lower and 'size' in source.content.lower()[:500]:
                    if not intel.market_size or len(stat) > len(intel.market_size):
                        intel.market_size = stat
                elif 'cagr' in stat_lower:
                    if not intel.market_cagr:
                        intel.market_cagr = stat
                elif 'margin' in stat_lower:
                    intel.industry_margins['benchmark'] = stat
        
        # LLM Synthesis for deeper insights
        if valid_sources:
            synthesis = await self.synthesize_research(valid_sources, sector)
            
            if synthesis:
                if synthesis.get('market_size') and synthesis['market_size'] != 'null':
                    intel.market_size = synthesis['market_size']
                if synthesis.get('market_cagr') and synthesis['market_cagr'] != 'null':
                    intel.market_cagr = synthesis['market_cagr']
                if synthesis.get('key_trends'):
                    intel.trends = [t for t in synthesis['key_trends'] if t][:4]
                if synthesis.get('growth_drivers'):
                    intel.growth_drivers = [d for d in synthesis['growth_drivers'] if d][:4]
                if synthesis.get('key_statistics'):
                    intel.statistics.update(synthesis['key_statistics'])
                if synthesis.get('investment_implications'):
                    intel.investment_implications = synthesis['investment_implications'][:3]
                if synthesis.get('competitive_landscape'):
                    # Extract company names from competitive landscape
                    comp_text = synthesis['competitive_landscape']
                    # Simple pattern to find capitalized company names
                    companies = re.findall(r'\b([A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)*)\b', comp_text)
                    intel.key_players = [c for c in companies if len(c) > 3][:5]
        
        # Generate executive summary
        if intel.market_size or intel.market_cagr or intel.trends:
            intel.executive_summary = await self._generate_executive_summary(intel, sector)
        
        print(f"  ✓ Research complete: Market={intel.market_size or 'N/A'}, "
              f"CAGR={intel.market_cagr or 'N/A'}, "
              f"{len(intel.trends)} trends, {len(intel.statistics)} stats")
        
        return intel
    
    async def _generate_executive_summary(self, intel: MarketIntelligence, 
                                          sector: str) -> str:
        """Generate a concise executive summary of market intelligence"""
        
        data_points = []
        if intel.market_size:
            data_points.append(f"Market size: {intel.market_size}")
        if intel.market_cagr:
            data_points.append(f"Growth: {intel.market_cagr}")
        if intel.trends:
            data_points.append(f"Key trends: {', '.join(intel.trends[:2])}")
        
        if not data_points:
            return ""
        
        prompt = f"""Write a 2-sentence executive summary for investors about the {sector} market.

DATA:
{chr(10).join(data_points)}

Write in professional investment banking style. Include specific numbers.
SUMMARY:"""

        return await self._call_llm(prompt, max_tokens=150, temperature=0.3)


# Convenience function for pipeline integration
async def perform_deep_research(sector: str, sub_sector: str = "", 
                                company_context: str = "") -> MarketIntelligence:
    """Perform deep web research and return market intelligence"""
    engine = AdvancedResearchEngine()
    try:
        return await engine.deep_research(sector, sub_sector, company_context)
    finally:
        await engine.close()
