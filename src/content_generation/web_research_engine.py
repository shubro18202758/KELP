"""
Web Research Engine
====================
Searches the web for real statistics, market data, and industry insights
to enrich investment teaser content with verified data.

Uses DuckDuckGo search + content extraction (no API keys needed).
"""

import re
import json
import asyncio
import aiohttp
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
from urllib.parse import quote_plus
import time


@dataclass
class ResearchResult:
    """Container for a single research finding"""
    query: str
    source: str
    url: str
    snippet: str
    extracted_numbers: List[str] = field(default_factory=list)
    relevance_score: float = 0.0


@dataclass
class MarketResearch:
    """Aggregated market research data"""
    market_size: str = ""
    market_cagr: str = ""
    key_players: List[str] = field(default_factory=list)
    industry_trends: List[str] = field(default_factory=list)
    growth_drivers: List[str] = field(default_factory=list)
    statistics: Dict[str, str] = field(default_factory=dict)
    sources: List[str] = field(default_factory=list)


class WebResearchEngine:
    """
    Searches the web for industry data, market statistics, and trends.
    
    Uses DuckDuckGo's HTML interface (no API key required) and extracts
    relevant statistics to enrich PPT content.
    """
    
    def __init__(self, ollama_base_url: str = "http://localhost:11434"):
        self.ollama_url = ollama_base_url
        self.model = "qwen2.5:7b"
        self.session = None
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        
    async def _get_session(self):
        """Get or create aiohttp session"""
        if self.session is None or self.session.closed:
            timeout = aiohttp.ClientTimeout(total=30)
            self.session = aiohttp.ClientSession(timeout=timeout, headers=self.headers)
        return self.session
    
    async def close(self):
        """Close the session"""
        if self.session and not self.session.closed:
            await self.session.close()
    
    async def search_duckduckgo(self, query: str, num_results: int = 8) -> List[Dict]:
        """
        Search DuckDuckGo and extract results.
        Returns list of {title, url, snippet}
        """
        results = []
        
        try:
            session = await self._get_session()
            
            # DuckDuckGo HTML search
            search_url = f"https://html.duckduckgo.com/html/?q={quote_plus(query)}"
            
            async with session.get(search_url) as resp:
                if resp.status == 200:
                    html = await resp.text()
                    
                    # Parse results from HTML
                    # DuckDuckGo uses specific classes for results
                    result_pattern = r'<a class="result__a" href="([^"]+)"[^>]*>([^<]+)</a>.*?<a class="result__snippet"[^>]*>([^<]+)</a>'
                    
                    # Simpler pattern for snippets
                    link_pattern = r'<a[^>]+class="result__a"[^>]+href="([^"]+)"[^>]*>([^<]+)</a>'
                    snippet_pattern = r'class="result__snippet"[^>]*>([^<]+)<'
                    
                    links = re.findall(link_pattern, html)
                    snippets = re.findall(snippet_pattern, html)
                    
                    for i, (url, title) in enumerate(links[:num_results]):
                        snippet = snippets[i] if i < len(snippets) else ""
                        
                        # Clean up URL (DuckDuckGo redirects)
                        if 'uddg=' in url:
                            actual_url = re.search(r'uddg=([^&]+)', url)
                            if actual_url:
                                from urllib.parse import unquote
                                url = unquote(actual_url.group(1))
                        
                        results.append({
                            'title': title.strip(),
                            'url': url,
                            'snippet': self._clean_text(snippet)
                        })
                        
        except Exception as e:
            print(f"  ⚠ DuckDuckGo search error: {e}")
        
        return results
    
    def _clean_text(self, text: str) -> str:
        """Clean HTML entities and extra whitespace"""
        text = re.sub(r'&[a-zA-Z]+;', ' ', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def _extract_numbers(self, text: str) -> List[str]:
        """Extract meaningful numbers and statistics from text"""
        patterns = [
            r'\$[\d,]+(?:\.\d+)?(?:\s*(?:billion|million|trillion|B|M|T|Cr|cr|Bn|bn))?',  # Currency
            r'₹[\d,]+(?:\.\d+)?(?:\s*(?:crore|lakh|Cr|L))?',  # Indian currency
            r'\d+(?:\.\d+)?%',  # Percentages
            r'(?:USD|INR|EUR)\s*[\d,]+(?:\.\d+)?(?:\s*(?:billion|million|B|M|Cr))?',
            r'\d+(?:\.\d+)?\s*(?:billion|million|trillion|crore|lakh)',  # Large numbers
            r'CAGR\s*(?:of\s*)?[\d.]+%',  # CAGR mentions
            r'(?:FY|Q[1-4])\s*\d{2,4}',  # Fiscal years/quarters
        ]
        
        numbers = []
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            numbers.extend(matches)
        
        return list(set(numbers))[:10]
    
    async def research_market_size(self, sector: str, sub_sector: str = "") -> Dict[str, str]:
        """
        Research market size and CAGR for the sector.
        Returns dict with market_size, cagr, and source.
        """
        query = f"{sector} {sub_sector} market size 2024 2025 India billion CAGR"
        results = await self.search_duckduckgo(query, 10)
        
        market_data = {
            'market_size': '',
            'cagr': '',
            'year': '',
            'source': ''
        }
        
        for result in results:
            text = f"{result['title']} {result['snippet']}"
            
            # Look for market size
            if not market_data['market_size']:
                size_match = re.search(
                    r'(?:market\s+size|valued\s+at|worth|estimated\s+at)[^\d]*'
                    r'((?:\$|USD|₹|INR)?\s*[\d.,]+\s*(?:billion|million|trillion|crore|Bn|B|M|Cr))',
                    text, re.IGNORECASE
                )
                if size_match:
                    market_data['market_size'] = size_match.group(1).strip()
                    market_data['source'] = result['url']
            
            # Look for CAGR
            if not market_data['cagr']:
                cagr_match = re.search(
                    r'CAGR\s*(?:of\s*)?([\d.]+\s*%)',
                    text, re.IGNORECASE
                )
                if cagr_match:
                    market_data['cagr'] = cagr_match.group(1)
        
        return market_data
    
    async def research_industry_trends(self, sector: str) -> List[str]:
        """Research current industry trends and growth drivers"""
        query = f"{sector} industry trends 2025 growth drivers India"
        results = await self.search_duckduckgo(query, 8)
        
        trends = []
        for result in results:
            snippet = result['snippet']
            if len(snippet) > 50:
                # Clean and truncate
                trend = snippet[:200] + "..." if len(snippet) > 200 else snippet
                trends.append(trend)
        
        return trends[:5]
    
    async def research_key_players(self, sector: str, sub_sector: str = "") -> List[str]:
        """Find key players/competitors in the sector"""
        query = f"{sector} {sub_sector} top companies players India market leaders"
        results = await self.search_duckduckgo(query, 6)
        
        players = []
        company_pattern = r'(?:top|leading|major)\s+(?:players?|companies?).*?(?:include|are|:)?\s*([A-Z][a-zA-Z]+(?:\s*,\s*[A-Z][a-zA-Z]+)*)'
        
        for result in results:
            text = result['snippet']
            matches = re.findall(company_pattern, text, re.IGNORECASE)
            for match in matches:
                companies = [c.strip() for c in match.split(',')]
                players.extend(companies)
        
        # Deduplicate and return top 5
        return list(dict.fromkeys(players))[:5]
    
    async def research_financial_benchmarks(self, sector: str) -> Dict[str, str]:
        """Research industry financial benchmarks (margins, multiples)"""
        query = f"{sector} India industry average EBITDA margin profit margin valuation multiples"
        results = await self.search_duckduckgo(query, 8)
        
        benchmarks = {}
        
        for result in results:
            text = result['snippet']
            
            # Look for margin data
            margin_match = re.search(r'(?:EBITDA|operating|profit)\s*margin[^\d]*([\d.]+\s*%)', text, re.IGNORECASE)
            if margin_match and 'ebitda_margin' not in benchmarks:
                benchmarks['industry_ebitda_margin'] = margin_match.group(1)
            
            # Look for growth rates
            growth_match = re.search(r'(?:revenue|sales)\s*growth[^\d]*([\d.]+\s*%)', text, re.IGNORECASE)
            if growth_match and 'revenue_growth' not in benchmarks:
                benchmarks['industry_growth'] = growth_match.group(1)
        
        return benchmarks
    
    async def _summarize_with_llm(self, research_data: List[Dict], sector: str) -> str:
        """Use LLM to summarize research findings into investor-ready insights"""
        combined_text = "\n".join([
            f"Source: {r['title']}\n{r['snippet']}" 
            for r in research_data[:6]
        ])
        
        prompt = f"""You are an M&A analyst summarizing market research for an investment teaser.

SECTOR: {sector}
RESEARCH DATA:
{combined_text}

Extract and summarize the KEY STATISTICS and INSIGHTS for investors:
1. Market size (with currency and year)
2. Growth rate / CAGR
3. Key industry trends
4. Major growth drivers

Format as a concise paragraph with specific numbers. Be factual, cite statistics.
If data is unclear, state the range or general trend.

SUMMARY (3-4 sentences with numbers):"""

        try:
            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=60)) as session:
                payload = {
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {"temperature": 0.3, "num_predict": 500}
                }
                async with session.post(f"{self.ollama_url}/api/generate", json=payload) as resp:
                    if resp.status == 200:
                        data = await resp.json()
                        return data.get("response", "").strip()
        except Exception as e:
            print(f"  ⚠ LLM summarization error: {e}")
        
        return ""
    
    async def comprehensive_research(self, sector: str, sub_sector: str = "", 
                                     company_name: str = "") -> MarketResearch:
        """
        Perform comprehensive web research for a sector.
        Returns aggregated market research data.
        """
        print(f"  🔍 Researching: {sector} / {sub_sector}")
        
        research = MarketResearch()
        
        # Parallel research tasks
        tasks = [
            self.research_market_size(sector, sub_sector),
            self.research_industry_trends(sector),
            self.research_key_players(sector, sub_sector),
            self.research_financial_benchmarks(sector)
        ]
        
        results = await asyncio.gather(*tasks, return_exceptions=True)
        
        # Process market size
        if isinstance(results[0], dict):
            market_data = results[0]
            if market_data.get('market_size'):
                research.market_size = market_data['market_size']
                research.statistics['Market Size'] = market_data['market_size']
            if market_data.get('cagr'):
                research.market_cagr = market_data['cagr']
                research.statistics['Industry CAGR'] = market_data['cagr']
            if market_data.get('source'):
                research.sources.append(market_data['source'])
        
        # Process trends
        if isinstance(results[1], list):
            research.industry_trends = results[1]
        
        # Process key players
        if isinstance(results[2], list):
            research.key_players = results[2]
        
        # Process benchmarks
        if isinstance(results[3], dict):
            research.statistics.update(results[3])
        
        # Generate growth drivers from trends
        if research.industry_trends:
            research.growth_drivers = [
                trend[:100] for trend in research.industry_trends[:3]
            ]
        
        print(f"  ✓ Found: Market size={research.market_size or 'N/A'}, "
              f"CAGR={research.market_cagr or 'N/A'}, "
              f"{len(research.statistics)} statistics")
        
        return research


class EnhancedContentGenerator:
    """
    Enhanced content generator that combines LLM with web research.
    """
    
    def __init__(self, ollama_url: str = "http://localhost:11434"):
        self.ollama_url = ollama_url
        self.model = "qwen2.5:7b"
        self.research_engine = WebResearchEngine(ollama_url)
    
    async def generate_investor_content(self, raw_data: str, sector: str, 
                                        sub_sector: str = "") -> Dict[str, Any]:
        """
        Generate investor-grade content enriched with web research.
        """
        # First, perform web research
        market_research = await self.research_engine.comprehensive_research(
            sector, sub_sector
        )
        
        # Combine company data with market research
        context = f"""
COMPANY DATA:
{raw_data[:4000]}

MARKET RESEARCH:
- Market Size: {market_research.market_size or 'Research pending'}
- Industry CAGR: {market_research.market_cagr or 'Research pending'}
- Key Players: {', '.join(market_research.key_players[:5]) or 'Various'}
- Industry Trends: {'; '.join(market_research.industry_trends[:3]) or 'Growing'}
"""
        
        # Generate enhanced content with market context
        content = await self._generate_with_context(context, sector)
        
        # Add research data
        content['market_research'] = {
            'market_size': market_research.market_size,
            'market_cagr': market_research.market_cagr,
            'statistics': market_research.statistics,
            'key_players': market_research.key_players,
            'sources': market_research.sources
        }
        
        return content
    
    async def _generate_with_context(self, context: str, sector: str) -> Dict[str, Any]:
        """Generate content using LLM with research context"""
        
        prompt = f"""You are a senior M&A investment banker at Goldman Sachs preparing a confidential 
investment teaser for a {sector} company. You have access to both company data and market research.

{context}

Generate the following in JSON format:

1. "executive_summary": 2-3 powerful sentences positioning this as a compelling investment opportunity
2. "investment_highlights": Array of 5 objects with "title" and "description" - each highlight should 
   reference specific metrics/statistics where available
3. "growth_story": Array of 4 compelling growth drivers with specific numbers
4. "market_positioning": 1-2 sentences about market position vs competitors
5. "key_statistics": Object with 4-6 key metrics that would impress investors

IMPORTANT:
- Use "The Company" or "Target" - never mention actual company name
- Include specific numbers, percentages, and market data where available
- Sound confident but factual
- Investment banker professional tone

OUTPUT (valid JSON only):"""

        try:
            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=90)) as session:
                payload = {
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": 0.7,
                        "num_predict": 2000,
                        "top_p": 0.9,
                    }
                }
                async with session.post(f"{self.ollama_url}/api/generate", json=payload) as resp:
                    if resp.status == 200:
                        data = await resp.json()
                        response = data.get("response", "")
                        
                        # Parse JSON from response
                        json_match = re.search(r'\{.*\}', response, re.DOTALL)
                        if json_match:
                            return json.loads(json_match.group())
                            
        except Exception as e:
            print(f"  ⚠ Enhanced generation error: {e}")
        
        # Fallback
        return {
            "executive_summary": f"Leading {sector} company with strong market position.",
            "investment_highlights": [],
            "growth_story": [],
            "market_positioning": "Well-positioned in growing market.",
            "key_statistics": {}
        }
    
    async def close(self):
        """Cleanup resources"""
        await self.research_engine.close()


# Test function
async def test_research():
    """Test the web research engine"""
    engine = WebResearchEngine()
    
    print("\n🔍 Testing Web Research Engine\n")
    
    # Test market size research
    print("1. Researching IT Services market...")
    market = await engine.research_market_size("IT Services", "Software Development")
    print(f"   Market Size: {market.get('market_size', 'Not found')}")
    print(f"   CAGR: {market.get('cagr', 'Not found')}")
    
    # Test comprehensive research
    print("\n2. Comprehensive research for Pharmaceuticals...")
    research = await engine.comprehensive_research("Pharmaceuticals", "Formulations")
    print(f"   Market Size: {research.market_size}")
    print(f"   Industry CAGR: {research.market_cagr}")
    print(f"   Key Players: {research.key_players[:3]}")
    print(f"   Statistics: {research.statistics}")
    
    await engine.close()
    print("\n✓ Research test complete!")


if __name__ == "__main__":
    asyncio.run(test_research())
