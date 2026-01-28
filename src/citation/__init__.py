"""
Citation Module - Generates citation documents for investment teasers
"""
from .citation_generator import (
    Citation,
    CitationCollection,
    CitationTracker,
    CitationDocumentGenerator,
    generate_citations_from_content
)

__all__ = [
    'Citation',
    'CitationCollection',
    'CitationTracker',
    'CitationDocumentGenerator',
    'generate_citations_from_content'
]
