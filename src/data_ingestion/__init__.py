"""
Data Ingestion Module
"""
from .markdown_parser import (
    CompanyData,
    MarkdownParser,
    CompanyDataExtractor,
    load_company_data,
    load_all_companies
)

__all__ = [
    'CompanyData',
    'MarkdownParser', 
    'CompanyDataExtractor',
    'load_company_data',
    'load_all_companies'
]
