"""
Sector Intelligence Module - Classifies companies and selects appropriate templates
"""
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from config.settings import SECTOR_CONFIGS, SectorConfig
from src.data_ingestion import CompanyData


@dataclass
class SectorClassification:
    """Result of sector classification"""
    sector_key: str
    sector_name: str
    confidence: float
    matched_keywords: List[str]
    config: SectorConfig


class SectorClassifier:
    """Classifies companies into sectors based on their data"""
    
    def __init__(self):
        self.sector_configs = SECTOR_CONFIGS
        
    def classify(self, company_data: CompanyData) -> SectorClassification:
        """
        Classify a company into a sector based on its data.
        Uses keyword matching and template hints.
        """
        scores: Dict[str, Tuple[float, List[str]]] = {}
        
        # Build searchable text from company data
        search_text = self._build_search_text(company_data)
        search_text_lower = search_text.lower()
        
        # Also check source file path for sector hints (folder name is reliable)
        source_file_lower = company_data.source_file.lower() if company_data.source_file else ""
        
        # Score each sector
        for sector_key, config in self.sector_configs.items():
            score, matched = self._score_sector(search_text_lower, config)
            
            # PRIORITY 1: Boost heavily if folder name contains sector key
            # This is the most reliable signal since folders are named by sector
            if sector_key in source_file_lower:
                score += 100  # Strong boost for folder name match
                matched.append(f"folder:{sector_key}")
            
            # PRIORITY 2: Boost if template matches (less reliable than folder)
            template_lower = company_data.sector_template.lower()
            if sector_key in template_lower or config.name.lower() in template_lower:
                score += 30  # Reduced from 50 since template can be wrong
                matched.append(f"template:{company_data.sector_template}")
            
            scores[sector_key] = (score, matched)
        
        # Find best match
        best_sector = max(scores.keys(), key=lambda k: scores[k][0])
        best_score, best_matched = scores[best_sector]
        
        # Calculate confidence (normalize to 0-1)
        max_possible = 150 + 30  # max keyword matches + folder + template bonus
        confidence = min(best_score / max_possible, 1.0)
        
        return SectorClassification(
            sector_key=best_sector,
            sector_name=self.sector_configs[best_sector].name,
            confidence=confidence,
            matched_keywords=best_matched,
            config=self.sector_configs[best_sector]
        )
    
    def _build_search_text(self, company_data: CompanyData) -> str:
        """Build searchable text from company data"""
        parts = [
            company_data.name,
            company_data.description,
            company_data.sector_template,
            company_data.domain,
            company_data.segment,
            ' '.join(company_data.products_services),
            ' '.join(company_data.industries_served),
            ' '.join(company_data.clients) if company_data.clients else ''
        ]
        return ' '.join(parts)
    
    def _score_sector(self, text: str, config: SectorConfig) -> Tuple[float, List[str]]:
        """Score how well text matches a sector"""
        score = 0
        matched = []
        
        for keyword in config.keywords:
            keyword_lower = keyword.lower()
            # Count occurrences
            count = text.count(keyword_lower)
            if count > 0:
                # Diminishing returns for repeated matches
                points = min(count * 5, 20)
                score += points
                matched.append(keyword)
        
        return score, matched
    
    def get_sector_config(self, sector_key: str) -> Optional[SectorConfig]:
        """Get configuration for a specific sector"""
        return self.sector_configs.get(sector_key)


class SlideContentSelector:
    """Selects and organizes content for each slide based on sector"""
    
    def __init__(self, sector_config: SectorConfig, company_data: CompanyData):
        self.config = sector_config
        self.data = company_data
        
    def get_slide1_content(self) -> Dict:
        """Get content for Slide 1: Business Profile"""
        content = {
            "title": self.config.slide1_title,
            "company_description": self._anonymize_description(self.data.description),
            "sections": {}
        }
        
        # Products/Services
        if self.data.products_services:
            content["sections"]["Products & Services"] = self.data.products_services[:8]
        
        # Industries Served
        if self.data.industries_served:
            content["sections"]["Industries Served"] = self.data.industries_served[:8]
        
        # Key Operational Highlights
        if self.data.key_operational_indicators:
            content["sections"]["Key Highlights"] = self.data.key_operational_indicators[:4]
        
        # Certifications (from awards)
        certifications = [a for a in self.data.awards_certifications 
                        if any(x in a.upper() for x in ['ISO', 'CMMI', 'GMP', 'IATF', 'CERT'])]
        if certifications:
            content["sections"]["Certifications"] = certifications[:6]
        
        # Global Presence
        if self.data.global_presence:
            content["sections"]["Global Presence"] = self.data.global_presence
        
        # Partners
        if self.data.partners:
            content["sections"]["Strategic Partners"] = self.data.partners
        
        return content
    
    def get_slide2_content(self) -> Dict:
        """Get content for Slide 2: Financial & Operational Metrics"""
        content = {
            "title": self.config.slide2_title,
            "metrics": [],
            "charts": []
        }
        
        # Extract metrics from SWOT Strengths (often contain financial data)
        if self.data.swot.get('Strengths'):
            for strength in self.data.swot['Strengths']:
                # Look for metrics with numbers
                if any(c.isdigit() for c in strength):
                    content["metrics"].append(self._extract_metric(strength))
        
        # Extract from operational indicators
        for indicator in self.data.key_operational_indicators:
            metric = self._extract_metric(indicator)
            if metric and metric not in content["metrics"]:
                content["metrics"].append(metric)
        
        # Channel mix data (for entertainment/retail)
        if self.data.channel_mix:
            content["channel_mix"] = self.data.channel_mix
        
        # Market size context
        if self.data.market_size_data:
            content["market_context"] = self.data.market_size_data[:2]
        
        # Prepare chart data
        content["charts"] = self._prepare_chart_data()
        
        return content
    
    def get_slide3_content(self) -> Dict:
        """Get content for Slide 3: Investment Highlights"""
        content = {
            "title": self.config.slide3_title,
            "highlights": [],
            "future_outlook": []
        }
        
        # Use sector-specific hooks as base
        content["highlights"] = self.config.slide3_hooks.copy()
        
        # Add from SWOT Opportunities
        if self.data.swot.get('Opportunities'):
            for opp in self.data.swot['Opportunities'][:2]:
                # Shorten if too long
                if len(opp) < 150:
                    content["highlights"].append(opp)
        
        # Future plans
        if self.data.future_plans:
            content["future_outlook"] = self.data.future_plans[:4]
        
        # Awards as credibility
        awards = [a for a in self.data.awards_certifications 
                 if 'award' in a.lower() or 'winner' in a.lower()]
        if awards:
            content["awards"] = awards[:3]
        
        return content
    
    def _anonymize_description(self, description: str) -> str:
        """Basic anonymization of company description"""
        if not description:
            return ""
        
        # This will be enhanced by LLM, but do basic cleanup
        text = description
        
        # Remove company name mentions (will be replaced by LLM)
        # For now, just return as-is - LLM will handle full anonymization
        return text
    
    def _extract_metric(self, text: str) -> Optional[str]:
        """Extract key metrics from text"""
        # Look for common metric patterns
        patterns = [
            r'(\d+(?:\.\d+)?%)',  # Percentages
            r'₹?\s*(\d+(?:,\d+)*(?:\.\d+)?)\s*(?:crore|cr|lakh|million|billion)',  # Currency
            r'(\d+)\s*(?:years?|months?)',  # Time periods
            r'(\d+)\+?\s*(?:employees?|customers?|clients?|screens?|cities?)',  # Counts
        ]
        
        for pattern in patterns:
            if re.search(pattern, text, re.IGNORECASE):
                # Clean up the text
                text = re.sub(r'\s+', ' ', text).strip()
                if len(text) < 200:
                    return text
                return text[:200] + "..."
        
        return None
    
    def _prepare_chart_data(self) -> List[Dict]:
        """Prepare data for charts based on available metrics"""
        charts = []
        
        # Try to extract growth metrics for bar chart
        growth_data = []
        for strength in self.data.swot.get('Strengths', []):
            # Look for percentage growth
            match = re.search(r'(\d+(?:\.\d+)?)\s*%\s*(?:growth|increase|grew|YoY)', strength, re.IGNORECASE)
            if match:
                growth_data.append({
                    "label": "Growth",
                    "value": float(match.group(1))
                })
        
        if growth_data:
            charts.append({
                "type": "bar",
                "title": "Key Growth Metrics",
                "data": growth_data
            })
        
        # Add placeholder for revenue/margin chart
        charts.append({
            "type": "bar",
            "title": "Financial Performance",
            "data": [
                {"label": "Revenue Growth", "value": 15},
                {"label": "Profit Growth", "value": 20},
                {"label": "Margin", "value": 12}
            ],
            "placeholder": True  # Mark as placeholder
        })
        
        return charts


def classify_company(company_data: CompanyData) -> Tuple[SectorClassification, Dict]:
    """
    Main function to classify a company and prepare slide content.
    Returns classification and organized content for all 3 slides.
    """
    classifier = SectorClassifier()
    classification = classifier.classify(company_data)
    
    selector = SlideContentSelector(classification.config, company_data)
    
    slide_content = {
        "slide1": selector.get_slide1_content(),
        "slide2": selector.get_slide2_content(),
        "slide3": selector.get_slide3_content()
    }
    
    return classification, slide_content


if __name__ == "__main__":
    # Test with sample data
    from src.data_ingestion import load_all_companies
    
    print("Testing Sector Classification...\n")
    companies = load_all_companies()
    
    for folder, company_data in companies.items():
        classification, content = classify_company(company_data)
        print(f"\n{'='*60}")
        print(f"Company: {company_data.name}")
        print(f"Classified as: {classification.sector_name}")
        print(f"Confidence: {classification.confidence:.2%}")
        print(f"Matched keywords: {', '.join(classification.matched_keywords[:5])}")
        print(f"Slide 1 sections: {len(content['slide1']['sections'])}")
        print(f"Slide 2 metrics: {len(content['slide2']['metrics'])}")
        print(f"Slide 3 highlights: {len(content['slide3']['highlights'])}")
