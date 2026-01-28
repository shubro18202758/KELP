"""
Investment-Grade Content Generator
===================================

Uses Ollama LLM (GPU-accelerated) to generate M&A investment banker quality content.
Transforms raw company data into professionally structured investment teaser content.

Key Features:
- Deep financial analysis extraction
- Investment thesis generation
- Growth story articulation
- Risk/opportunity assessment
- Anonymized "blind" content creation
"""

import re
import json
import asyncio
import aiohttp
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class InvestmentContent:
    """Structured content for investment teaser"""
    # Cover
    project_codename: str = ""
    sector_classification: str = ""
    sub_sector: str = ""
    
    # Business Overview
    executive_summary: str = ""
    business_description: List[str] = field(default_factory=list)
    value_proposition: str = ""
    
    # Products & Markets
    product_portfolio: List[str] = field(default_factory=list)
    target_industries: List[str] = field(default_factory=list)
    key_customers: List[str] = field(default_factory=list)
    competitive_position: str = ""
    
    # Financial Metrics
    revenue_metrics: Dict[str, Any] = field(default_factory=dict)
    profitability_metrics: Dict[str, Any] = field(default_factory=dict)
    growth_metrics: Dict[str, Any] = field(default_factory=dict)
    
    # Operational
    manufacturing_footprint: Dict[str, Any] = field(default_factory=dict)
    employee_count: int = 0
    certifications: List[str] = field(default_factory=list)
    
    # Investment Thesis
    investment_highlights: List[Dict[str, str]] = field(default_factory=list)
    growth_drivers: List[str] = field(default_factory=list)
    expansion_plans: List[str] = field(default_factory=list)
    
    # Risk Assessment
    key_risks: List[str] = field(default_factory=list)
    mitigants: List[str] = field(default_factory=list)


class InvestmentContentGenerator:
    """
    Generates investment-grade content using local LLM.
    
    Uses carefully crafted prompts to extract and structure
    M&A-quality content from raw company data.
    """
    
    def __init__(self, ollama_base_url: str = "http://localhost:11434"):
        self.base_url = ollama_base_url
        self.model = "qwen2.5:7b"
        self._available = None
        
    async def check_availability(self) -> bool:
        """Check if Ollama is running"""
        if self._available is not None:
            return self._available
            
        try:
            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=5)) as session:
                async with session.get(f"{self.base_url}/api/tags") as resp:
                    self._available = resp.status == 200
        except:
            self._available = False
        return self._available
    
    async def _generate(self, prompt: str, max_tokens: int = 2000, 
                        temperature: float = 0.4) -> str:
        """
        Generate content using LLM with optimized parameters.
        
        Lower temperature (0.3-0.5) produces more focused, factual content
        ideal for investment presentations where precision matters.
        """
        if not await self.check_availability():
            return ""
            
        try:
            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=120)) as session:
                payload = {
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": temperature,  # Lower for factual precision
                        "num_predict": max_tokens,
                        "num_gpu": 99,  # Use all available GPU layers
                        "top_p": 0.85,  # Slightly tighter for coherence
                        "repeat_penalty": 1.15,  # Avoid repetition
                        "num_ctx": 4096,  # Larger context window
                    }
                }
                async with session.post(f"{self.base_url}/api/generate", json=payload) as resp:
                    if resp.status == 200:
                        data = await resp.json()
                        return data.get("response", "")
        except Exception as e:
            print(f"  ⚠ LLM generation error: {e}")
        return ""
    
    async def generate_business_overview(self, raw_data: str, sector: str) -> List[str]:
        """
        Generate investment banker quality business overview bullets.
        
        Returns 5-6 key bullets that would appear in a teaser.
        """
        prompt = f"""You are a senior M&A investment banker at Goldman Sachs preparing a CONFIDENTIAL 
investment teaser for a $500M+ deal. This document will be read by institutional investors, 
PE funds, and strategic acquirers who demand specific, quantifiable data points.

SECTOR: {sector}

COMPANY DATA:
{raw_data[:4000]}

CRITICAL REQUIREMENTS - Each bullet MUST include:
1. SPECIFIC NUMBERS: Revenue figures, market share %, CAGR, unit volumes, capacity utilization
2. COMPETITIVE POSITIONING: "Top 3 in...", "#1 provider of...", "Largest independent..."
3. SCALE INDICATORS: Employee count, facility count, customer count, geographic reach
4. PROVEN RESULTS: YoY growth %, margin improvement, contract values, retention rates

STYLE:
- Investment banking professional language (concise, powerful, data-driven)
- DO NOT mention company name - use "The Company" or "Target"
- Each bullet: 1-2 sentences with AT LEAST one specific metric
- Avoid vague language like "significant", "growing" - replace with NUMBERS

EXAMPLE STRONG BULLETS:
- "The Company commands ~18% market share in India's ₹12,000 Cr automotive forging sector"
- "Revenue grew at 22% CAGR from FY21-FY24, reaching ₹850 Cr with 14.5% EBITDA margin"
- "Operating 5 manufacturing units with combined 45,000 MT capacity at 78% utilization"

Format: Return as a JSON array of 5-6 strings.
OUTPUT (JSON array only):"""

        response = await self._generate(prompt, 1200)
        
        try:
            # Parse JSON response
            json_match = re.search(r'\[.*\]', response, re.DOTALL)
            if json_match:
                bullets = json.loads(json_match.group())
                return [str(b).strip() for b in bullets if b]
        except:
            pass
        
        # Fallback: split by newlines
        bullets = [line.strip().strip('- •"') for line in response.split('\n') if line.strip()]
        return bullets[:6]
    
    async def generate_investment_highlights(self, raw_data: str, sector: str,
                                             financials: Dict) -> List[Dict[str, str]]:
        """
        Generate 5 compelling investment highlights.
        
        Each highlight has a bold title and supporting description.
        """
        fin_summary = json.dumps(financials, indent=2) if financials else "Not available"
        
        prompt = f"""You are a senior M&A advisor preparing investment highlights for PE fund principals 
and strategic acquirers. These are the 5 key points that will convince investors to pay a premium.

SECTOR: {sector}
COMPANY DATA:
{raw_data[:3500]}

FINANCIALS:
{fin_summary}

EACH HIGHLIGHT MUST HAVE:
1. TITLE (10-15 words): Start with action verb or strong positioning statement
   - Use specific numbers: "%", "₹ Cr", "#1", "Top 5"
   - Examples: "18% Revenue CAGR with Consistent 400+ bps Margin Expansion"
   
2. DESCRIPTION (2 sentences): Elaborate with supporting evidence
   - Include additional metrics, customer names, contract terms
   - Reference competitive advantages and market dynamics

HIGHLIGHT CATEGORIES TO COVER:
- Market Leadership: Share, ranking, competitive moat
- Financial Excellence: CAGR, margins, ROCE, debt profile  
- Operational Scale: Capacity, facilities, workforce, technology
- Growth Runway: TAM expansion, new products, geographic expansion
- Strategic Value: Synergy potential, platform acquisition thesis

DO NOT mention company name - use "The Company" or "Target"

Return as JSON array:
[
  {{"title": "Market Leader with 22% Share in ₹8,000 Cr Addressable Market", 
    "description": "Ranked #1 in precision forging with 35-year operating history. Serving 8 of top 10 OEMs with 95% repeat customer rate."}},
  ...
]

OUTPUT (JSON array with 5 highlights):"""

        response = await self._generate(prompt, 1800)
        
        try:
            json_match = re.search(r'\[.*\]', response, re.DOTALL)
            if json_match:
                highlights = json.loads(json_match.group())
                return [h for h in highlights if isinstance(h, dict) and 'title' in h]
        except:
            pass
        
        # Fallback defaults
        return [
            {"title": f"Leading player in {sector} with proprietary capabilities",
             "description": "Strong market position with differentiated offerings"},
            {"title": "Serving blue-chip client base with high retention",
             "description": "Long-standing relationships with marquee customers"},
            {"title": "Strategically located manufacturing infrastructure",
             "description": "Optimal footprint for cost efficiency and market access"},
            {"title": "Strong financial profile with consistent growth",
             "description": "Demonstrated track record of profitable growth"},
            {"title": "Significant expansion potential",
             "description": "Multiple levers for future growth identified"}
        ]
    
    async def generate_growth_story(self, raw_data: str, sector: str) -> List[str]:
        """
        Generate 3-4 growth story callouts for financial slide.
        """
        prompt = f"""You are creating the "Growth Drivers" section for an investment teaser's 
financial slide. These bullet points justify the premium valuation investors will pay.

SECTOR: {sector}
COMPANY DATA:
{raw_data[:3000]}

EACH GROWTH DRIVER MUST:
1. Be SPECIFIC and QUANTIFIABLE - include %s, ₹ figures, timelines
2. Reference concrete initiatives (not vague "growth opportunities")
3. Show cause-and-effect logic: "X initiative → Y% improvement expected"

TYPES OF GROWTH DRIVERS TO IDENTIFY:
- Capacity expansion: "New ₹150 Cr facility adding 40% capacity by Q2 FY26"
- Margin improvement: "Product mix shift to drive 300bps EBITDA expansion over 2 years"
- Market expansion: "Entry into electric vehicle segment, targeting ₹100 Cr by FY27"
- Operating leverage: "Fixed cost absorption improving as utilization hits 85%+"
- Pricing power: "Annual price escalation clauses with top 5 customers (avg 3-4%)"

AVOID GENERIC STATEMENTS. Write like a Goldman Sachs MD justifying a 12x EBITDA multiple.

Return as JSON array of 4 strings with specific numbers.
OUTPUT:"""

        response = await self._generate(prompt, 1000)
        
        try:
            json_match = re.search(r'\[.*\]', response, re.DOTALL)
            if json_match:
                return json.loads(json_match.group())[:4]
        except:
            pass
        
        return [
            "Strong margin expansion driven by product mix improvement",
            "Capacity expansion adding 25%+ to existing infrastructure",
            "Secured long-term contracts ensuring revenue visibility"
        ]
    
    async def generate_upcoming_facility(self, raw_data: str) -> List[str]:
        """
        Generate upcoming facility/expansion bullet points.
        """
        prompt = f"""Extract expansion plans and upcoming facility information from the data.
Create 4 bullet points about planned investments/expansion.

DATA:
{raw_data[:2500]}

REQUIREMENTS:
1. Each bullet: specific about capex, capacity, timeline, expected returns
2. If data not available, create plausible points for the industry
3. Use checkmark style points (✓)

Return as JSON array of strings (without checkmarks).
OUTPUT:"""

        response = await self._generate(prompt, 600)
        
        try:
            json_match = re.search(r'\[.*\]', response, re.DOTALL)
            if json_match:
                return json.loads(json_match.group())[:4]
        except:
            pass
        
        return [
            "Planned capex for capacity expansion",
            "New facility to add incremental revenue",
            "Expected commissioning within 18 months",
            "Superior margins expected from new capacity"
        ]
    
    async def anonymize_content(self, text: str, entities_to_remove: List[str] = None) -> str:
        """
        Anonymize content by removing company-specific identifiers.
        """
        if not text:
            return text
        
        prompt = f"""Anonymize this text for a blind investment teaser.

ORIGINAL TEXT:
{text}

REQUIREMENTS:
1. Replace company name with "The Company" or "Target"
2. Replace specific locations with generic terms ("strategic location", "key manufacturing hub")
3. Keep all numbers, metrics, and financial data intact
4. Maintain professional investment banking tone
5. Do NOT change the meaning or facts

ANONYMIZED TEXT:"""

        response = await self._generate(prompt, 500)
        return response.strip() if response else text
    
    async def generate_full_teaser_content(self, raw_data: str, sector: str,
                                            financials: Dict = None) -> InvestmentContent:
        """
        Generate complete investment teaser content.
        
        Orchestrates all content generation for a full teaser.
        """
        content = InvestmentContent()
        content.sector_classification = sector
        
        print("  🚀 Generating business overview with GPU...")
        content.business_description = await self.generate_business_overview(raw_data, sector)
        
        print("  🚀 Generating investment highlights with GPU...")
        content.investment_highlights = await self.generate_investment_highlights(
            raw_data, sector, financials or {}
        )
        
        print("  🚀 Generating growth story with GPU...")
        content.growth_drivers = await self.generate_growth_story(raw_data, sector)
        
        print("  🚀 Generating expansion plans with GPU...")
        content.expansion_plans = await self.generate_upcoming_facility(raw_data)
        
        return content
    

class SmartWebResearchGuide:
    """
    Uses LLM to guide web research for better data collection.
    """
    
    def __init__(self, ollama_base_url: str = "http://localhost:11434"):
        self.base_url = ollama_base_url
        self.model = "qwen2.5:7b"
    
    async def generate_search_queries(self, company_name: str, sector: str) -> List[str]:
        """
        Generate intelligent search queries for web research.
        """
        prompt = f"""Generate 5 search queries to find investment-relevant information about a company.

COMPANY: {company_name}
SECTOR: {sector}

Generate queries for:
1. Recent news and announcements
2. Financial performance and growth
3. Products and competitive position
4. Awards and certifications
5. Future plans and expansion

Return as JSON array of search query strings.
OUTPUT:"""

        try:
            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=30)) as session:
                payload = {
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {"temperature": 0.3, "num_predict": 400}
                }
                async with session.post(f"{self.base_url}/api/generate", json=payload) as resp:
                    if resp.status == 200:
                        data = await resp.json()
                        response = data.get("response", "")
                        
                        json_match = re.search(r'\[.*\]', response, re.DOTALL)
                        if json_match:
                            return json.loads(json_match.group())[:5]
        except:
            pass
        
        # Fallback queries
        return [
            f"{company_name} financial performance revenue growth",
            f"{company_name} products services market position",
            f"{company_name} news expansion plans 2025",
            f"{company_name} awards certifications quality",
            f"{company_name} {sector} industry analysis"
        ]
    
    async def extract_key_facts(self, web_content: str, sector: str) -> List[str]:
        """
        Extract key investment-relevant facts from web content.
        """
        prompt = f"""Extract key investment-relevant facts from this web content.

SECTOR: {sector}
CONTENT:
{web_content[:4000]}

Extract:
1. Revenue/financial metrics
2. Market position statements
3. Growth achievements
4. Customer/client mentions
5. Awards/recognition

Return as JSON array of fact strings (max 10 facts).
OUTPUT:"""

        try:
            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=30)) as session:
                payload = {
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {"temperature": 0.2, "num_predict": 600}
                }
                async with session.post(f"{self.base_url}/api/generate", json=payload) as resp:
                    if resp.status == 200:
                        data = await resp.json()
                        response = data.get("response", "")
                        
                        json_match = re.search(r'\[.*\]', response, re.DOTALL)
                        if json_match:
                            return json.loads(json_match.group())[:10]
        except:
            pass
        
        return []


# ============================================================================
# INTEGRATED CONTENT PIPELINE
# ============================================================================

async def generate_teaser_content_gpu(raw_markdown: str, sector: str,
                                       financials: Dict = None,
                                       verbose: bool = True) -> Dict:
    """
    Main entry point for GPU-accelerated content generation.
    
    Enhanced with web research data integration for real market statistics.
    
    Args:
        raw_markdown: Raw company data from markdown files
        sector: Classified sector
        financials: Dict containing financial metrics + optional web research:
            - revenue, ebitda_margin, employees (core financials)
            - market_size, market_cagr (from web research)
            - industry_trends, key_players, growth_drivers (from web research)
        verbose: Whether to print progress
    
    Returns:
        Dictionary ready for PPT generation with investment-grade content.
    """
    generator = InvestmentContentGenerator()
    
    if not await generator.check_availability():
        if verbose:
            print("  ⚠ GPU LLM not available, using fallback content")
        return _generate_fallback_content(sector)
    
    if verbose:
        print("  🚀 GPU-accelerated content generation starting...")
    
    # Enhance raw data with web research if available
    enhanced_context = raw_markdown
    if financials:
        market_intelligence = []
        if financials.get('market_size'):
            market_intelligence.append(f"MARKET SIZE: {financials['market_size']}")
        if financials.get('market_cagr'):
            market_intelligence.append(f"INDUSTRY CAGR: {financials['market_cagr']}")
        if financials.get('industry_trends'):
            trends = financials['industry_trends']
            market_intelligence.append(f"KEY TRENDS: {'; '.join(trends[:3])}")
        if financials.get('key_players'):
            players = financials['key_players']
            market_intelligence.append(f"COMPETITIVE LANDSCAPE: {', '.join(players[:5])}")
        if financials.get('growth_drivers'):
            drivers = financials['growth_drivers']
            market_intelligence.append(f"MARKET GROWTH DRIVERS: {'; '.join(drivers[:3])}")
        
        if market_intelligence:
            enhanced_context = f"""[WEB RESEARCH - USE THESE STATISTICS IN YOUR CONTENT]
{chr(10).join(market_intelligence)}

[COMPANY DATA]
{raw_markdown}"""
            if verbose:
                print("  📊 Enriched with web-researched market intelligence")
    
    content = await generator.generate_full_teaser_content(enhanced_context, sector, financials)
    
    # Transform to PPT-ready format
    return {
        'sector': sector,
        'slide1': {
            'company_description': content.business_description[0] if content.business_description else '',
            'key_points': content.business_description[1:],
        },
        'slide2': {
            'financial_highlights': [],
            'growth_drivers': content.growth_drivers,
        },
        'slide3': {
            'investment_highlights': content.investment_highlights,
        },
        'business_overview': content.business_description,
        'investment_highlights': content.investment_highlights,
        'growth_highlights': content.growth_drivers,
        'upcoming_facility': content.expansion_plans,
        # Include market research in output for teaser data
        'market_intelligence': {
            'market_size': financials.get('market_size') if financials else None,
            'market_cagr': financials.get('market_cagr') if financials else None,
            'trends': financials.get('industry_trends') if financials else [],
        }
    }


def _generate_fallback_content(sector: str) -> Dict:
    """Generate fallback content when LLM is not available"""
    sector_defaults = {
        "Manufacturing": {
            'business_overview': [
                "Leading manufacturer with integrated operations and strong market position",
                "Multiple state-of-the-art manufacturing facilities across strategic locations",
                "Serving diverse industries with customized product solutions",
                "Strong focus on quality with international certifications",
                "Experienced management team with deep industry expertise"
            ],
            'investment_highlights': [
                {"title": "Market leader with proprietary manufacturing capabilities",
                 "description": "Strong competitive position with high entry barriers"},
                {"title": "Blue-chip customer base with high retention",
                 "description": "Long-term relationships with marquee clients"},
                {"title": "Strategic manufacturing footprint",
                 "description": "Optimally located facilities for cost efficiency"},
                {"title": "Strong financial profile",
                 "description": "Consistent growth with healthy margins"},
                {"title": "Significant expansion potential",
                 "description": "Multiple growth levers identified"}
            ]
        },
        "Technology": {
            'business_overview': [
                "Technology solutions provider with specialized domain expertise",
                "Global delivery capabilities with presence in key markets",
                "Strong engineering talent pool and R&D capabilities",
                "Recurring revenue model with long-term client relationships",
                "Track record of successful digital transformation projects"
            ],
            'investment_highlights': [
                {"title": "Deep technology expertise in high-growth domains",
                 "description": "Specialized capabilities in emerging technology areas"},
                {"title": "Sticky client relationships with Fortune 500 companies",
                 "description": "High client retention and expanding wallet share"},
                {"title": "Scalable delivery model",
                 "description": "Proven ability to scale operations efficiently"},
                {"title": "Strong margin profile",
                 "description": "Industry-leading profitability metrics"},
                {"title": "Multiple growth vectors",
                 "description": "Geographic and service line expansion opportunities"}
            ]
        }
    }
    
    defaults = sector_defaults.get(sector, sector_defaults["Manufacturing"])
    
    return {
        'sector': sector,
        'slide1': {
            'company_description': defaults['business_overview'][0],
            'key_points': defaults['business_overview'][1:],
        },
        'slide2': {
            'growth_drivers': [
                "Strong margin expansion through operational improvements",
                "Capacity addition to capture growing demand",
                "New customer acquisitions in target markets"
            ],
        },
        'slide3': {
            'investment_highlights': defaults['investment_highlights'],
        },
        'business_overview': defaults['business_overview'],
        'investment_highlights': defaults['investment_highlights'],
        'growth_highlights': [
            "Margin expansion through efficiency improvements",
            "Capacity addition to meet growing demand"
        ],
        'upcoming_facility': [
            "Planned capacity expansion",
            "New facility commissioning expected",
            "Incremental revenue opportunity"
        ],
    }


# ============================================================================
# CLI Test
# ============================================================================

if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).parent.parent.parent))
    
    async def test():
        print("=" * 60)
        print("INVESTMENT CONTENT GENERATOR - TEST")
        print("=" * 60)
        
        generator = InvestmentContentGenerator()
        
        if await generator.check_availability():
            print("✓ GPU LLM available")
            
            sample_data = """
            Kalyani Forge Ltd. is a leading Indian engineering company specializing in 
            high-quality forged, machined, and assembled products for automotive, 
            construction, and power generation industries. The company has 5 manufacturing 
            plants with ISO/TS 16949 certification. Revenue of ₹2,366 Cr with 10% EBITDA margin.
            Key clients include JCB, Daimler, Mahindra, Honda, Cummins.
            """
            
            print("\nGenerating business overview...")
            bullets = await generator.generate_business_overview(sample_data, "Manufacturing")
            for b in bullets:
                print(f"  • {b}")
            
            print("\nGenerating investment highlights...")
            highlights = await generator.generate_investment_highlights(
                sample_data, "Manufacturing", {"revenue": "2366 Cr", "ebitda_margin": "10%"}
            )
            for h in highlights:
                print(f"  📌 {h.get('title', '')}")
                print(f"     {h.get('description', '')}")
        else:
            print("✗ GPU LLM not available")
    
    asyncio.run(test())
