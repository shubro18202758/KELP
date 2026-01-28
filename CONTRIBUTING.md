# Contributing to KELP

Thank you for your interest in contributing to KELP! This document provides guidelines for contributing to the project.

## 🚀 Getting Started

### Prerequisites

1. Python 3.10 or higher
2. NVIDIA GPU with CUDA support (recommended)
3. Ollama installed locally
4. Git

### Development Setup

```bash
# Clone the repository
git clone https://github.com/shubro18202758/KELP.git
cd KELP

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Linux/Mac
.\venv\Scripts\Activate   # Windows

# Install dependencies
pip install -r requirements.txt
pip install ddgs

# Install development dependencies
pip install pytest black flake8 mypy

# Setup Ollama
ollama pull qwen2.5:7b
```

## 📋 How to Contribute

### Reporting Bugs

1. Check existing issues to avoid duplicates
2. Use the bug report template
3. Include:
   - Python version
   - OS and version
   - GPU model and VRAM
   - Full error traceback
   - Steps to reproduce

### Suggesting Features

1. Check existing feature requests
2. Describe the problem the feature solves
3. Propose a solution
4. Consider implementation complexity

### Pull Requests

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Make your changes
4. Run tests: `python main.py --company kalyani`
5. Format code: `black .`
6. Commit with descriptive message
7. Push to your fork
8. Open a Pull Request

## 🎨 Code Style

### Python Style Guide

- Follow PEP 8
- Use type hints where possible
- Maximum line length: 100 characters
- Use descriptive variable names

```python
# Good
def generate_investment_teaser(company_data: CompanyData) -> TeaserOutput:
    """Generate an investment teaser from company data."""
    pass

# Bad
def gen_teaser(d):
    pass
```

### Docstrings

Use Google-style docstrings:

```python
def process_company(
    company_path: Path,
    output_dir: Path,
    verbose: bool = False
) -> PipelineResult:
    """Process a single company through the pipeline.
    
    Args:
        company_path: Path to company data directory.
        output_dir: Directory for output files.
        verbose: Enable verbose logging.
        
    Returns:
        PipelineResult containing status and output paths.
        
    Raises:
        ValueError: If company data is invalid.
    """
    pass
```

### Commit Messages

Use conventional commits:

```
feat: add new chart type for financial slides
fix: resolve web search timeout issue
docs: update installation instructions
refactor: simplify content generator logic
test: add unit tests for markdown parser
```

## 🏗️ Project Structure

```
KELP/
├── src/                    # Source code
│   ├── content_generation/ # LLM content generation
│   ├── data_ingestion/     # Data parsing
│   ├── presentation/       # PPT generation
│   └── ...
├── config/                 # Configuration
├── docs/                   # Documentation
├── tests/                  # Test files
└── Company Data/           # Sample data
```

### Adding New Features

1. **New chart type**: Modify `src/presentation/enhanced_kelp_generator.py`
2. **New sector**: Update `src/sector_intelligence/classifier.py`
3. **New LLM prompt**: Edit `src/content_generation/investment_content_generator.py`
4. **New data source**: Add to `src/web_scraping/`

## 🧪 Testing

```bash
# Run on test company
python main.py --company kalyani

# Verify output exists
ls output/v5_enhanced/

# Check for errors
python main.py --company kalyani 2>&1 | grep -i error
```

## 📝 Documentation

- Update README.md for user-facing changes
- Add docstrings to new functions
- Update config/settings.py comments for new settings

## 🤝 Code Review

All PRs require:
- At least one approval
- Passing tests
- No merge conflicts
- Clean commit history

## 📜 License

By contributing, you agree that your contributions will be licensed under the MIT License.

---

Thank you for contributing to KELP! 🌊
