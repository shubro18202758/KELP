"""
Web Scraping Module
Contains intelligent web scraper with DuckDuckGo search integration
"""
from .scraper import (
    WebScrapedData,
    AsyncWebScraper,
    IntelligentWebScraper,
    scrape_company_website
)

from .web_search import (
    SearchResult,
    ExtractedContent,
    CompanyWebData,
    DuckDuckGoSearch,
    IntelligentScraper,
    WebSearchPipeline,
    research_company,
    research_company_async
)

__all__ = [
    # From scraper.py
    'WebScrapedData',
    'AsyncWebScraper', 
    'IntelligentWebScraper',
    'scrape_company_website',
    # From web_search.py
    'SearchResult',
    'ExtractedContent',
    'CompanyWebData',
    'DuckDuckGoSearch',
    'IntelligentScraper',
    'WebSearchPipeline',
    'research_company',
    'research_company_async'
]
