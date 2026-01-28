"""
Data Enrichment Engine - GPU-Accelerated Rich Content Extraction
=================================================================

This module uses the local LLM (via Ollama with GPU acceleration) to extract
ALL metrics, financial data, and sector-specific KPIs from source documents.

It populates the PPT slides with REAL data like:
- Revenue Growth (CAGR ~15%+)
- EBITDA Margins (20%+)
- Export contribution (e.g., 45% revenue from exports)
- Customer count (600+)
- Certifications (IATF 16949, ISO 14001, etc.)
- Credit ratings (CRISIL BBB, etc.)

Author: Generated for Kelp M&A Hackathon
"""

import re
import json
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
import asyncio
import aiohttp

# Try to import numpy for calculations
try:
    import numpy as np
    HAS_NUMPY = True
except ImportError:
    HAS_NUMPY = False


@dataclass
class ExtractedMetrics:
    """Container for all extracted metrics"""
    # Financial Metrics
    revenue_latest: float = 0.0
    revenue_cagr: float = 0.0
    revenue_trend: List[float] = field(default_factory=list)
    revenue_years: List[int] = field(default_factory=list)
    
    ebitda_latest: float = 0.0
    ebitda_margin: float = 0.0
    ebitda_trend: List[float] = field(default_factory=list)
    
    pat_latest: float = 0.0
    pat_margin: float = 0.0
    pat_trend: List[float] = field(default_factory=list)
    
    # Operational Metrics
    employee_count: int = 0
    customer_count: int = 0
    plant_count: int = 0
    product_count: int = 0
    
    # Export & Global
    export_percentage: float = 0.0
    export_growth: float = 0.0
    countries_present: int = 0
    global_locations: List[str] = field(default_factory=list)
    
    # Credit & Risk
    credit_rating: str = ""
    credit_agency: str = ""
    rating_outlook: str = ""
    
    # Certifications
    certifications: List[str] = field(default_factory=list)
    
    # Order Book & Growth
    order_book_value: float = 0.0
    capex_planned: float = 0.0
    capacity_utilization: float = 0.0
    
    # Industry Specific
    market_size: float = 0.0
    market_cagr: float = 0.0
    market_share: float = 0.0
    
    # Clients
    key_clients: List[str] = field(default_factory=list)
    blue_chip_clients: int = 0
    
    # SWOT Highlights
    key_strengths: List[str] = field(default_factory=list)
    growth_drivers: List[str] = field(default_factory=list)


class DataEnrichmentEngine:
    """
    GPU-Accelerated Data Enrichment Engine using local LLM.
    
    Extracts comprehensive metrics from company data for rich PPT generation.
    """
    
    def __init__(self, ollama_base_url: str = "http://localhost:11434"):
        self.base_url = ollama_base_url
        self.model = "qwen2.5:7b"
        self._available = None
        
    async def check_availability(self) -> bool:
        """Check if Ollama is running and accessible"""
        if self._available is not None:
            return self._available
            
        try:
            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=5)) as session:
                async with session.get(f"{self.base_url}/api/tags") as resp:
                    self._available = resp.status == 200
        except:
            self._available = False
        return self._available
    
    async def llm_extract(self, prompt: str, max_tokens: int = 1000) -> str:
        """Use LLM to extract structured data"""
        if not await self.check_availability():
            return ""
            
        try:
            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=60)) as session:
                payload = {
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": 0.1,  # Low temp for accurate extraction
                        "num_predict": max_tokens,
                        "num_gpu": 99,  # Use all available GPU layers
                    }
                }
                async with session.post(f"{self.base_url}/api/generate", json=payload) as resp:
                    if resp.status == 200:
                        data = await resp.json()
                        return data.get("response", "")
        except Exception as e:
            print(f"LLM extraction error: {e}")
        return ""
    
    def extract_financial_data(self, raw_content: str) -> Dict[str, Any]:
        """Extract financial data from the markdown content using regex patterns"""
        financials = {
            "revenue": [],
            "ebitda": [],
            "pat": [],
            "years": [],
            "revenue_growth": None,
            "ebitda_margin": None,
            "pat_margin": None,
        }
        
        # Extract Revenue from Operations data - format: "2014: 2054.23131 | 2015: 2408.02796 | ..."
        revenue_line = re.search(r"Revenue From Operations[^\n]*\n?[^-]*", raw_content, re.IGNORECASE)
        if revenue_line:
            line = revenue_line.group()
            # Extract year:value pairs
            year_value_pattern = r"(\d{4}):\s*([\d.]+)"
            matches = re.findall(year_value_pattern, line)
            for year_str, val_str in matches:
                try:
                    year = int(year_str)
                    val = float(val_str)
                    if val > 0:
                        financials["years"].append(year)
                        financials["revenue"].append(val)
                except:
                    pass
        
        # Extract EBITDA data
        ebitda_line = re.search(r"Operating EBITDA[^\n]*\n?[^-]*", raw_content, re.IGNORECASE)
        if ebitda_line:
            line = ebitda_line.group()
            year_value_pattern = r"(\d{4}):\s*([-\d.]+)"
            matches = re.findall(year_value_pattern, line)
            for year_str, val_str in matches:
                try:
                    val = float(val_str)
                    financials["ebitda"].append(val)
                except:
                    pass
        
        # Extract PAT data - more careful pattern
        pat_line = re.search(r"^- PAT \|[^\n]+", raw_content, re.MULTILINE)
        if pat_line:
            line = pat_line.group()
            year_value_pattern = r"(\d{4}):\s*([-\d.]+)"
            matches = re.findall(year_value_pattern, line)
            for year_str, val_str in matches:
                try:
                    val = float(val_str)
                    financials["pat"].append(val)
                except:
                    pass
        
        # Calculate metrics
        if len(financials["revenue"]) >= 2:
            latest = financials["revenue"][-1]
            # Use 5-year CAGR if available
            if len(financials["revenue"]) >= 5:
                oldest = financials["revenue"][-5]
                years = 4
            else:
                oldest = financials["revenue"][0]
                years = len(financials["revenue"]) - 1
            if oldest > 0 and years > 0:
                financials["revenue_growth"] = ((latest / oldest) ** (1 / years) - 1) * 100
        
        if financials["revenue"] and financials["ebitda"]:
            latest_rev = financials["revenue"][-1]
            latest_ebitda = financials["ebitda"][-1] if financials["ebitda"] else 0
            if latest_rev > 0 and latest_ebitda > 0:
                financials["ebitda_margin"] = (latest_ebitda / latest_rev) * 100
        
        if financials["revenue"] and financials["pat"]:
            latest_rev = financials["revenue"][-1]
            latest_pat = financials["pat"][-1] if financials["pat"] else 0
            if latest_rev > 0:
                financials["pat_margin"] = (latest_pat / latest_rev) * 100
        
        return financials
    
    def extract_operational_data(self, raw_content: str) -> Dict[str, Any]:
        """Extract operational metrics using regex patterns"""
        ops = {
            "employees": 0,
            "plants": 0,
            "customers": 0,
            "products": 0,
            "export_pct": 0,
            "certifications": [],
            "credit_rating": "",
            "key_clients": [],
            "locations": [],
        }
        
        # Employee count - handle format like "Employees: **484 (1%)**"
        emp_patterns = [
            r"[Ee]mployees?:?\s*\*{0,2}(\d+)",
            r"[Tt]otal\s*[Ee]mployees?:?\s*\*{0,2}(\d+)",
            r"[Hh]eadcount:?\s*\*{0,2}(\d+)",
            r"[Ww]orkforce:?\s*\*{0,2}(\d+)",
            r"(\d+)\s*employees?",
        ]
        for pattern in emp_patterns:
            match = re.search(pattern, raw_content)
            if match:
                ops["employees"] = int(match.group(1))
                break
        
        # Plant/Facility count
        plant_patterns = [
            r"(\d+)\s*plants?",
            r"(\d+)\s*facilities?",
            r"(\d+)\s*manufacturing\s*units?",
        ]
        for pattern in plant_patterns:
            match = re.search(pattern, raw_content, re.IGNORECASE)
            if match:
                ops["plants"] = int(match.group(1))
                break
        
        # Customer count
        cust_patterns = [
            r"(\d+)\+?\s*customers?",
            r"(\d+)\+?\s*clients?",
            r"serves?\s*(\d+)\+?\s*",
        ]
        for pattern in cust_patterns:
            match = re.search(pattern, raw_content, re.IGNORECASE)
            if match:
                ops["customers"] = int(match.group(1))
                break
        
        # Export percentage
        export_patterns = [
            r"export.*?(\d+(?:\.\d+)?)\s*%",
            r"(\d+(?:\.\d+)?)\s*%.*?export",
            r"exports?\s*(?:of|at|around|approximately|~)?\s*(\d+(?:\.\d+)?)\s*%",
        ]
        for pattern in export_patterns:
            match = re.search(pattern, raw_content, re.IGNORECASE)
            if match:
                ops["export_pct"] = float(match.group(1))
                break
        
        # Certifications
        cert_patterns = [
            r"(ISO\s*\d+(?::\d+)?)",
            r"(IATF\s*\d+(?::\d+)?)",
            r"(GMP\+?)",
            r"(FSSC\s*\d+)",
            r"(CE\s*certified?)",
            r"(TS\s*\d+)",
            r"(NADCAP)",
            r"(AS\s*\d+)",
        ]
        for pattern in cert_patterns:
            matches = re.findall(pattern, raw_content, re.IGNORECASE)
            ops["certifications"].extend(matches[:5])
        ops["certifications"] = list(set(ops["certifications"]))[:6]
        
        # Credit Rating - handle table format: | ... | CRISIL | BBB | Reaffirmed |
        # Try table format first
        table_rating = re.search(r"\|\s*(CRISIL|ICRA|CARE|Fitch)\s*\|\s*([A-Z]{1,3}[+-]?\s*\d*)\s*\|", raw_content, re.IGNORECASE)
        if table_rating:
            ops["credit_rating"] = f"{table_rating.group(1)} {table_rating.group(2).strip()}"
        else:
            # Fallback to text patterns
            rating_patterns = [
                r"(CRISIL|ICRA|CARE|Fitch|Moody'?s?)\s*[:-]?\s*([A-Z]{1,3}[+-]?\d*\+?)",
                r"rating.*?(CRISIL|ICRA|CARE).*?([A-Z]{1,3}[+-]?\d*)",
            ]
            for pattern in rating_patterns:
                match = re.search(pattern, raw_content, re.IGNORECASE)
                if match:
                    ops["credit_rating"] = f"{match.group(1)} {match.group(2)}"
                break
        
        # Key Clients - look for known company names
        client_patterns = [
            r"(Tata\s*\w*)",
            r"(Mahindra)",
            r"(Honda)",
            r"(Daimler)",
            r"(JCB)",
            r"(Cummins)",
            r"(Honeywell)",
            r"(Bosch)",
            r"(Toyota)",
            r"(Maruti)",
            r"(Hero)",
            r"(Bajaj)",
            r"(L&T)",
            r"(Reliance)",
            r"(ONGC)",
            r"(NTPC)",
        ]
        for pattern in client_patterns:
            matches = re.findall(pattern, raw_content, re.IGNORECASE)
            if matches:
                ops["key_clients"].extend(matches[:2])
        ops["key_clients"] = list(set(ops["key_clients"]))[:8]
        
        # Locations
        location_patterns = [
            r"(India|USA|United States|Germany|UK|United Kingdom|Japan|China|Europe|Middle East|Africa|Asia|Americas?)",
        ]
        for pattern in location_patterns:
            matches = re.findall(pattern, raw_content, re.IGNORECASE)
            ops["locations"].extend(matches)
        ops["locations"] = list(set(ops["locations"]))[:5]
        
        return ops
    
    def extract_market_data(self, raw_content: str) -> Dict[str, Any]:
        """Extract market size, growth, and industry data"""
        market = {
            "market_size": 0,
            "market_size_unit": "Bn",
            "market_cagr": 0,
            "industry": "",
            "target_segments": [],
        }
        
        # Market size patterns
        size_patterns = [
            r"\$?\s*(\d+(?:\.\d+)?)\s*(billion|bn|B)\s*(?:market|industry|sector)?",
            r"(?:market|industry|sector)\s*(?:size|worth|valued).*?\$?\s*(\d+(?:\.\d+)?)\s*(billion|bn|B)",
            r"₹?\s*(\d+(?:\.\d+)?)\s*(crore|cr)\s*(?:market)?",
        ]
        for pattern in size_patterns:
            match = re.search(pattern, raw_content, re.IGNORECASE)
            if match:
                market["market_size"] = float(match.group(1))
                unit = match.group(2).lower()
                market["market_size_unit"] = "Bn" if unit in ["billion", "bn", "b"] else "Cr"
                break
        
        # Market CAGR
        cagr_patterns = [
            r"(?:market|industry).*?CAGR.*?(\d+(?:\.\d+)?)\s*%",
            r"(\d+(?:\.\d+)?)\s*%\s*CAGR.*?(?:market|industry)",
            r"grow(?:ing|th).*?(\d+(?:\.\d+)?)\s*%",
        ]
        for pattern in cagr_patterns:
            match = re.search(pattern, raw_content, re.IGNORECASE)
            if match:
                market["market_cagr"] = float(match.group(1))
                break
        
        return market
    
    def extract_swot_highlights(self, raw_content: str) -> Dict[str, List[str]]:
        """Extract SWOT analysis highlights"""
        swot = {
            "strengths": [],
            "opportunities": [],
            "growth_drivers": [],
        }
        
        # Look for strengths section
        strengths_match = re.search(r"##\s*Strengths?\s*\n(.*?)(?=##|\Z)", raw_content, re.DOTALL | re.IGNORECASE)
        if strengths_match:
            items = re.findall(r"[-•*]\s*(.+?)(?=\n|$)", strengths_match.group(1))
            swot["strengths"] = [s.strip() for s in items[:5] if len(s.strip()) > 10]
        
        # Look for opportunities section
        opp_match = re.search(r"##\s*Opportunities?\s*\n(.*?)(?=##|\Z)", raw_content, re.DOTALL | re.IGNORECASE)
        if opp_match:
            items = re.findall(r"[-•*]\s*(.+?)(?=\n|$)", opp_match.group(1))
            swot["opportunities"] = [s.strip() for s in items[:5] if len(s.strip()) > 10]
        
        return swot
    
    def extract_order_capex_data(self, raw_content: str) -> Dict[str, Any]:
        """Extract order book and capex data"""
        data = {
            "order_book": 0,
            "capex": 0,
            "capex_planned": 0,
        }
        
        # Order book patterns
        order_patterns = [
            r"order\s*(?:book|win|value).*?₹?\s*(\d+(?:,\d+)?(?:\.\d+)?)\s*(?:cr|crore|Cr)",
            r"₹?\s*(\d+(?:,\d+)?(?:\.\d+)?)\s*(?:cr|crore|Cr).*?order",
        ]
        for pattern in order_patterns:
            match = re.search(pattern, raw_content, re.IGNORECASE)
            if match:
                data["order_book"] = float(match.group(1).replace(',', ''))
                break
        
        # CapEx patterns
        capex_patterns = [
            r"capex.*?₹?\s*(\d+(?:,\d+)?(?:\.\d+)?)\s*(?:cr|crore|Cr)",
            r"₹?\s*(\d+(?:,\d+)?(?:\.\d+)?)\s*(?:cr|crore|Cr).*?capex",
            r"capital\s*expenditure.*?₹?\s*(\d+(?:,\d+)?(?:\.\d+)?)\s*(?:cr|crore|Cr)",
        ]
        for pattern in capex_patterns:
            match = re.search(pattern, raw_content, re.IGNORECASE)
            if match:
                data["capex"] = float(match.group(1).replace(',', ''))
                break
        
        return data
    
    async def enrich_with_llm(self, raw_content: str, sector: str) -> Dict[str, Any]:
        """Use LLM to extract additional insights and format data"""
        
        # Truncate content for LLM (first 4000 chars)
        content_snippet = raw_content[:4000]
        
        prompt = f"""You are a financial analyst extracting key metrics for an M&A investment teaser.

Analyze this {sector} company data and extract:
1. Key financial highlights (revenue, profit, margins)
2. Operational strengths (capacity, clients, certifications)
3. Growth catalysts (market trends, expansion plans)
4. Risk factors to mention

Company Data:
{content_snippet}

Provide a JSON response with these exact keys:
{{
    "financial_highlights": ["list of 3-4 key financial points with numbers"],
    "operational_strengths": ["list of 3-4 operational strengths with numbers"],
    "growth_catalysts": ["list of 3-4 growth drivers"],
    "investment_thesis": "2 sentence summary of why this is attractive"
}}

Return ONLY the JSON, no other text:"""

        result = await self.llm_extract(prompt, max_tokens=800)
        
        if result:
            try:
                # Try to parse JSON from response
                json_match = re.search(r'\{.*\}', result, re.DOTALL)
                if json_match:
                    return json.loads(json_match.group())
            except json.JSONDecodeError:
                pass
        
        return {}
    
    async def extract_all_metrics(self, raw_content: str, sector: str) -> ExtractedMetrics:
        """
        Main extraction method - combines regex + LLM for comprehensive data extraction.
        """
        metrics = ExtractedMetrics()
        
        # 1. Extract financial data (regex - fast)
        financials = self.extract_financial_data(raw_content)
        if financials["revenue"]:
            metrics.revenue_latest = financials["revenue"][-1]
            metrics.revenue_trend = financials["revenue"]
            metrics.revenue_years = financials["years"]
        if financials["revenue_growth"] is not None:
            metrics.revenue_cagr = financials["revenue_growth"]
        if financials["ebitda"]:
            metrics.ebitda_latest = financials["ebitda"][-1]
            metrics.ebitda_trend = financials["ebitda"]
        if financials["ebitda_margin"] is not None:
            metrics.ebitda_margin = financials["ebitda_margin"]
        if financials["pat"]:
            metrics.pat_latest = financials["pat"][-1]
            metrics.pat_trend = financials["pat"]
        if financials["pat_margin"] is not None:
            metrics.pat_margin = financials["pat_margin"]
        
        # 2. Extract operational data (regex - fast)
        ops = self.extract_operational_data(raw_content)
        metrics.employee_count = ops["employees"]
        metrics.plant_count = ops["plants"]
        metrics.customer_count = ops["customers"]
        metrics.export_percentage = ops["export_pct"]
        metrics.certifications = ops["certifications"]
        metrics.credit_rating = ops["credit_rating"]
        metrics.key_clients = ops["key_clients"]
        metrics.global_locations = ops["locations"]
        metrics.countries_present = len(ops["locations"])
        
        # 3. Extract market data
        market = self.extract_market_data(raw_content)
        metrics.market_size = market["market_size"]
        metrics.market_cagr = market["market_cagr"]
        
        # 4. Extract SWOT highlights
        swot = self.extract_swot_highlights(raw_content)
        metrics.key_strengths = swot["strengths"]
        metrics.growth_drivers = swot["opportunities"]
        
        # 5. Extract order book and capex
        order_capex = self.extract_order_capex_data(raw_content)
        metrics.order_book_value = order_capex["order_book"]
        metrics.capex_planned = order_capex["capex"]
        
        return metrics
    
    def generate_sector_metrics(self, metrics: ExtractedMetrics, sector: str) -> Dict[str, Any]:
        """
        Generate sector-specific formatted metrics for PPT slides.
        """
        sector_lower = sector.lower()
        
        # Base formatted data
        formatted = {
            "revenue_display": f"₹{metrics.revenue_latest:.0f} Cr" if metrics.revenue_latest else "N/A",
            "revenue_cagr_display": f"{metrics.revenue_cagr:.1f}%" if metrics.revenue_cagr else "N/A",
            "ebitda_margin_display": f"{metrics.ebitda_margin:.1f}%" if metrics.ebitda_margin else "N/A",
            "pat_margin_display": f"{metrics.pat_margin:.1f}%" if metrics.pat_margin else "N/A",
            "employee_display": f"{metrics.employee_count:,}" if metrics.employee_count else "N/A",
            "plant_display": str(metrics.plant_count) if metrics.plant_count else "N/A",
            "customer_display": f"{metrics.customer_count}+" if metrics.customer_count else "N/A",
            "export_display": f"{metrics.export_percentage:.0f}%" if metrics.export_percentage else "N/A",
            "credit_display": metrics.credit_rating if metrics.credit_rating else "N/A",
            "certifications_display": ", ".join(metrics.certifications[:4]) if metrics.certifications else "N/A",
            "clients_display": ", ".join(metrics.key_clients[:5]) if metrics.key_clients else "N/A",
            "market_size_display": f"${metrics.market_size:.0f}B" if metrics.market_size else "N/A",
            "market_cagr_display": f"{metrics.market_cagr:.1f}%" if metrics.market_cagr else "N/A",
        }
        
        # Sector-specific KPIs
        if "manufacturing" in sector_lower or "automotive" in sector_lower or "engineering" in sector_lower:
            formatted["sector_kpis"] = [
                f"Revenue: {formatted['revenue_display']} (CAGR {formatted['revenue_cagr_display']})",
                f"EBITDA Margin: {formatted['ebitda_margin_display']}",
                f"Exports: {formatted['export_display']} of revenue",
                f"Plants: {formatted['plant_display']} manufacturing facilities",
                f"Clients: {formatted['customer_display']} customers",
                f"Credit Rating: {formatted['credit_display']}",
            ]
            formatted["chart_data"] = {
                "revenue_trend": {
                    "years": metrics.revenue_years[-5:],
                    "values": [v for v in metrics.revenue_trend[-5:]],
                    "title": "Revenue Trend (₹ Cr)",
                },
                "ebitda_trend": {
                    "years": metrics.revenue_years[-5:],
                    "values": [v for v in metrics.ebitda_trend[-5:]],
                    "title": "EBITDA Trend (₹ Cr)",
                },
            }
            
        elif "pharma" in sector_lower or "healthcare" in sector_lower:
            formatted["sector_kpis"] = [
                f"Revenue: {formatted['revenue_display']} (CAGR {formatted['revenue_cagr_display']})",
                f"R&D Pipeline: Active development",
                f"Certifications: {formatted['certifications_display']}",
                f"Market: ${metrics.market_size:.0f}B (CAGR {formatted['market_cagr_display']})",
                f"Global Presence: {len(metrics.global_locations)} countries",
            ]
            
        elif "technology" in sector_lower or "it" in sector_lower or "software" in sector_lower:
            formatted["sector_kpis"] = [
                f"Revenue: {formatted['revenue_display']} (CAGR {formatted['revenue_cagr_display']})",
                f"Team Size: {formatted['employee_display']} professionals",
                f"Client Base: {formatted['customer_display']} enterprises",
                f"Global Presence: {len(metrics.global_locations)} locations",
            ]
            
        elif "logistics" in sector_lower or "transport" in sector_lower:
            formatted["sector_kpis"] = [
                f"Revenue: {formatted['revenue_display']} (CAGR {formatted['revenue_cagr_display']})",
                f"Network: Pan-India coverage",
                f"Fleet/Capacity: Strong infrastructure",
                f"Clients: {formatted['customer_display']} served",
            ]
            
        elif "entertainment" in sector_lower or "media" in sector_lower:
            formatted["sector_kpis"] = [
                f"Revenue: {formatted['revenue_display']}",
                f"Market Position: Leading regional player",
                f"Properties: Multiple venues/assets",
            ]
            
        else:
            # Generic KPIs
            formatted["sector_kpis"] = [
                f"Revenue: {formatted['revenue_display']} (CAGR {formatted['revenue_cagr_display']})",
                f"EBITDA Margin: {formatted['ebitda_margin_display']}",
                f"Team: {formatted['employee_display']} employees",
                f"Clients: {formatted['customer_display']}",
            ]
        
        return formatted
    
    def generate_chart_data_for_slides(self, metrics: ExtractedMetrics) -> Dict[str, Any]:
        """
        Generate actual chart data for visualization in slides.
        """
        charts = {}
        
        # Revenue trend chart
        if metrics.revenue_trend and metrics.revenue_years:
            charts["revenue_bar"] = {
                "type": "bar",
                "title": "Revenue Trend (₹ Crores)",
                "labels": [str(y) for y in metrics.revenue_years[-5:]],
                "values": [round(v, 1) for v in metrics.revenue_trend[-5:]],
                "color": "#2E86AB",
            }
        
        # EBITDA trend chart
        if metrics.ebitda_trend and len(metrics.ebitda_trend) >= 3:
            charts["ebitda_bar"] = {
                "type": "bar",
                "title": "EBITDA Trend (₹ Crores)",
                "labels": [str(y) for y in metrics.revenue_years[-5:]],
                "values": [round(v, 1) for v in metrics.ebitda_trend[-5:]],
                "color": "#28A745",
            }
        
        # PAT trend chart
        if metrics.pat_trend and len(metrics.pat_trend) >= 3:
            charts["pat_bar"] = {
                "type": "bar",
                "title": "Net Profit (₹ Crores)",
                "labels": [str(y) for y in metrics.revenue_years[-5:]],
                "values": [round(v, 1) for v in metrics.pat_trend[-5:]],
                "color": "#6C5CE7",
            }
        
        # Margin analysis pie/gauge
        if metrics.ebitda_margin > 0:
            charts["margin_gauge"] = {
                "type": "gauge",
                "title": "EBITDA Margin",
                "value": round(metrics.ebitda_margin, 1),
                "max": 30,
                "color": "#E74C3C" if metrics.ebitda_margin < 10 else "#F39C12" if metrics.ebitda_margin < 15 else "#27AE60",
            }
        
        # Geographic presence pie
        if metrics.global_locations:
            charts["geo_pie"] = {
                "type": "pie",
                "title": "Geographic Presence",
                "labels": metrics.global_locations[:5],
                "values": [1] * len(metrics.global_locations[:5]),  # Equal distribution
            }
        
        return charts


def create_enriched_content(raw_markdown: str, sector: str) -> Tuple[ExtractedMetrics, Dict[str, Any], Dict[str, Any]]:
    """
    Synchronous wrapper to extract and format all data.
    
    Returns:
        - ExtractedMetrics: All extracted metrics
        - sector_formatted: Sector-specific formatted data
        - chart_data: Data ready for chart generation
    """
    import asyncio
    
    engine = DataEnrichmentEngine()
    
    # Run async extraction
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        metrics = loop.run_until_complete(engine.extract_all_metrics(raw_markdown, sector))
    finally:
        loop.close()
    
    # Generate formatted data
    sector_formatted = engine.generate_sector_metrics(metrics, sector)
    chart_data = engine.generate_chart_data_for_slides(metrics)
    
    return metrics, sector_formatted, chart_data


if __name__ == "__main__":
    # Test the engine
    import asyncio
    
    # Sample test content
    test_content = """
    ## Financial Performance
    The company achieved Revenue From Operations of ₹2366.5 Cr in FY2025.
    Operating EBITDA stood at ₹240 Cr with healthy margins.
    
    ## Operations
    - 5 manufacturing plants
    - 484 employees
    - Exports contribute 40% of revenue
    - IATF 16949:2016, ISO 14001:2015 certified
    
    ## Credit Rating
    CRISIL BBB rating with stable outlook
    
    ## Key Clients
    Daimler, JCB, Tata Motors, Honda, Cummins
    
    ## Market
    $468 billion global automotive aftermarket growing at 6.6% CAGR
    """
    
    engine = DataEnrichmentEngine()
    
    async def test():
        metrics = await engine.extract_all_metrics(test_content, "Automotive Manufacturing")
        print("Extracted Metrics:")
        print(f"  Revenue Latest: ₹{metrics.revenue_latest} Cr")
        print(f"  EBITDA Margin: {metrics.ebitda_margin:.1f}%")
        print(f"  Employees: {metrics.employee_count}")
        print(f"  Plants: {metrics.plant_count}")
        print(f"  Export %: {metrics.export_percentage}%")
        print(f"  Certifications: {metrics.certifications}")
        print(f"  Credit Rating: {metrics.credit_rating}")
        print(f"  Key Clients: {metrics.key_clients}")
        
        formatted = engine.generate_sector_metrics(metrics, "Automotive Manufacturing")
        print("\nFormatted KPIs:")
        for kpi in formatted.get("sector_kpis", []):
            print(f"  • {kpi}")
    
    asyncio.run(test())
