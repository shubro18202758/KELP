<img src="https://r2cdn.perplexity.ai/pplx-full-logo-primary-dark%402x.png" style="height:64px;margin-right:32px"/>

# okay since this is a whole institute wide hackathon at iitb and the problem statement is attached as pdf so I want you to go through it deeply and then provide me a plan of action as well as ideal model according to the problem statement and the provided use cases...also I will be using qwen3 8b model in rtx 4070 gpu (8GB variant) in legion 7i with i9 14900hx cpu instead of any api (and don't give any code yet I just want an overview)

Now let me provide you with the comprehensive strategic overview:

***

# 🎯 Kelp Automated Deal Flow - Strategic Hackathon Overview

## Executive Summary

You're tackling an **M\&A automation pipeline** that transforms a company's data pack into a professional 3-slide "blind" investment teaser. This is a **full-stack AI problem** combining data engineering, NLP, sector intelligence, and presentation automation. Given your hardware (RTX 4070 + Qwen 3 8B), you're well-positioned to solve this locally without API dependencies.

***

## 🏗️ The Complete System Architecture

### **Phase 1: Hybrid Data Ingestion (The Foundation)**

Your pipeline must ingest data from **two fundamentally different sources**:

**Private Data Layer** consists of Kelp-provided proprietary files—Excel balance sheets, PDF credit reports, historical M\&A data. This is structured, reliable, but specific to each company. The challenge here isn't finding the data; it's **parsing heterogeneous formats** (Excel sheets with varying column orders, PDFs with inconsistent layouts, potential OCR errors). Your data ingestion layer needs robustness: multiple parsers, error handling for malformed data, fallback strategies when a field is missing.

**Public Data Layer** pulls qualitative context—company websites, blog posts, LinkedIn profiles, news articles. This is where most real-world complexity lives: websites break, content is behind paywalls, some companies have minimal online presence. The strategic challenge is **intelligent filtering**—scraping a company website might return 500+ paragraphs, but only 10-20 are relevant to the investment thesis. Your scraper needs semantic understanding to isolate the signal.

**Why This Matters for Evaluation**: The "Data Fusion Capability" criterion (20% weight) explicitly judges how well you combine these two sources. A shallow approach treats them separately; a strong solution creates coherent narrative by linking private metrics with public positioning.

***

### **Phase 2: Sector Detection Engine (The Intelligence Layer)**

This is where **25% of your evaluation** happens, and it's the differentiator between a passing and winning solution.

The problem: A manufacturing company requires completely different metrics than a D2C e-commerce brand. The evaluation matrix explicitly penalizes generic slide templates. So your system needs to **classify the company's industry** from the provided data, then **select the right template** for that sector.

**Classification Strategy**:

- Parse company description, business model, product types, financial metrics
- Use heuristics + Qwen 3 8B for semantic classification
- Assign confidence scores
- Trigger sector-specific slide logic

**Sector-Specific Variations** (from problem statement):


| Sector | Slide 1 Focus | Slide 2 Focus | Slide 3 Hook |
| :-- | :-- | :-- | :-- |
| **Manufacturing** | Business profile, product segments, manufacturing footprint | Revenue growth (CAGR 15%+), EBITDA margins (20%+), certifications | Proprietary portfolio, entry barriers, logistics advantages |
| **D2C** | Brand overview, portfolio mix, channel presence | Growth \& unit economics (LTV/CAC > 7x), repeat rate (35%+), CAGR (65%) | Top-ranked positioning, profitability (70%+ gross margin), market whitespace |
| **Pharma** | Research pipeline, regulatory status, patent portfolio | Clinical outcomes, approval status, R\&D spend efficiency | Patent moat, market opportunity, regulatory advantages |
| **Tech** | Product value prop, user growth, TAM | Retention metrics, ARR growth, unit economics | Platform stickiness, network effects, TAM capture |
| **Logistics** | Network footprint, operational scale, geographic coverage | Throughput, cost per delivery, utilization efficiency | Consolidation play, capital efficiency, scale advantages |

The beauty here is that **you demonstrate sector expertise programmatically**. A generic "3-slide structure" loses this weightage; a smart sector classifier wins it.

***

### **Phase 3: Qwen 3 8B as Your Content Brain**

Given your hardware constraints and API budget (₹100/company), **Qwen 3 8B local inference is the optimal choice**. Here's why:

**Why Qwen 3 8B Works**:

1. **Sufficient Reasoning for Financial Analysis**
    - The model can parse Excel financials, understand KPIs, calculate growth rates
    - Can synthesize disparate data points into coherent narratives
    - Strong at numerical reasoning (CAGR calculations, margin analysis)
2. **Anonymization at Scale**
    - This is a creative task requiring both fact-preservation and identity-concealment
    - Qwen can rewrite text to eliminate company identifiers while maintaining accuracy
    - Can understand what constitutes "revealing" information (specific facility names, unique metric combinations)
3. **Hardware Fit**
    - RTX 4070 (8GB) can run Qwen 3 8B in fp16 (~5-6GB) or int8 quantization (~4GB)
    - Legion 7i's i9-14900HX provides strong CPU support for batch processing
    - Expected inference time: 1-2 minutes per slide of content generation
4. **Zero API Cost**
    - All inference happens locally—you stay well under ₹100 budget
    - Only real costs: image APIs (which have generous free tiers)

**How to Deploy**:

```
Option A: Ollama (Easiest)
- Install Ollama, pull Qwen model
- Simple REST API for inference
- ~30 seconds startup time

Option B: vLLM (Faster)
- More optimized inference
- Batch processing capabilities
- Steeper learning curve

Recommendation: Start with Ollama, migrate to vLLM if inference becomes bottleneck
```


***

### **Phase 4: Anonymization Engine (The Creative Challenge)**

This is **15% of your evaluation**, and it's surprisingly nuanced.

**The Problem**: You must make the teaser completely unidentifiable to a competitor (can't mention company name, no logos, no unique geographic markers) while keeping it convincing to an investor. If anonymization is clumsy ("Company XYZ manufactures specialty chemicals"), you lose points. If it's too vague ("There is a manufacturing company"), it becomes worthless.

**Anonymization Strategy**:

1. **Entity Stripping**: Remove names, founder identities, specific facility locations
2. **Semantic Generalization**:
    - ✗ "Apollo Tyres' Chennai facility produces..."
    - ✓ "A leading manufacturer of automotive components operates state-of-the-art production facilities..."
3. **Metric Transformation**:
    - Keep numbers accurate but contextualize differently
    - "Revenue of ₹500 crore" → "Mid-market player in ₹4000 crore category"
    - Avoids unique identification while preserving investor relevance
4. **Paraphrasing with LLM**:
    - Feed extracted facts to Qwen
    - Prompt: "Rewrite this company description to anonymize the company while preserving facts and improving presentation quality"
    - Few-shot examples help set the tone

**Key Insight**: The best anonymization doesn't feel anonymized—it reads like a genuine teaser that happens to omit the company name.

***

### **Phase 5: Image Intelligence (The Visual Layer)**

Every slide needs **3-4 high-quality, sector-appropriate images** that are:

- Generic enough not to reveal the company
- Directly relevant to the business type
- Professional enough for investor presentations

**Image Sourcing Strategy**:

1. **Query Generation**: Transform sector classification into image search queries
    - Manufacturing → "chemical facility", "production line", "R\&D laboratory"
    - D2C → "wellness products lifestyle", "e-commerce fulfillment", "product packaging"
    - Pharma → "pharmaceutical research lab", "clinical testing"
2. **API Strategy**:
    - Primary: Unsplash / Pexels (free, >100M images, high quality)
    - Fallback: Bing Image Search API (paid, but reliable filtering)
    - Cost: ~₹0-10 per company
3. **Validation**:
    - Programmatic check for visible logos/text
    - Size/resolution validation
    - Manual spot-check for 1-2 companies

***

### **Phase 6: Native PPT Generation (The Production Challenge)**

This is **30% of your evaluation**—the highest weighted component. You CANNOT generate PDF/images and embed them. The output must be a **native, editable PowerPoint file** with **actual chart objects**, not screenshots.

**Why This Matters**: If an investor receives your teaser, they want to edit it, change colors, update numbers. A static PDF is useless. A PNG mockup is worthless.

**Branding Requirements** (from problem statement):


| Component | Specification |
| :-- | :-- |
| **Logo** | "Kelp" text placeholder, top of every slide |
| **Footer** | "Strictly Private \& Confidential – Prepared by Kelp M\&A Team" (9pt, centered bottom) |
| **Primary Color** | Dark Indigo/Violet (covers \& overlays) |
| **Secondary** | Pink-to-Orange Gradient (accents), Cyan Blue (icons) |
| **Background** | Clean White (content slides) |
| **Text** | White (on covers), Dark Grey (body) |
| **Headings** | Arial Bold, 20-24pt |
| **Body Text** | Arial Regular / Aptos, 10-12pt |
| **Layout** | 3-4 quadrants/columns, no text walls, clean imagery |

**Technical Implementation** (`python-pptx` library):

```
For each of 3 slides:
1. Add title + subtitle with proper styling
2. Place 1-2 images (full bleed or grid layout)
3. Add content blocks (bullet points, metrics)
4. IF Slide 2: Generate native chart objects
   - Bar chart: Revenue CAGR + margins
   - Pie chart: Revenue mix / channel distribution
   - Line chart: Growth trajectory
5. Apply color palette programmatically
6. Add footer with compliance text
```

**Chart Generation**: Use `python-pptx`'s built-in chart API or generate with Plotly → export as PPT-compatible format. **Key risk**: Defaulting to embedding PNG charts (static images). Winners create native objects (editable, resizable).

***

### **Phase 7: Citation Document (The Accountability Layer)**

Every claim in your teaser must be **traceable to a source**. This is **10% of evaluation** and demonstrates rigor.

**Citation Document Structure**:

```
Slide 1 - Business Profile
├─ "Leading manufacturer of specialty chemicals" 
│  └─ Source: Company website (provided data pack)
├─ "150+ customers across paints, pharma, nutrition sectors"
│  └─ Source: Private data: CustomerDatabase.xlsx (Sheet: "Active Customers")
└─ [Image attribution]
   └─ Source: Unsplash / Pexels [URL]

Slide 2 - Financial Metrics
├─ "CAGR of 15% over 3 years"
│  └─ Source: Private data: BalanceSheet.xlsx (Rows 5-10, Column C)
├─ "EBITDA margin of 22%"
│  └─ Calculation: [Revenue - COGS - Operating Expenses] / Revenue (FY2024, BalanceSheet.xlsx)
└─ Chart data
   └─ Sources: BalanceSheet.xlsx, OperationalMetrics.xlsx

Slide 3 - Investment Highlights
├─ "Industry-leading product portfolio"
│  └─ Source: Company website + product catalog (provided data pack)
├─ "Strategic geographic advantage"
│  └─ Source: Derived from logistics analysis + facility locations in private data
└─ ...
```

**Implementation**: Metadata tracking during extraction phase—every data point tagged with source URL or file reference.

***

## 📊 Evaluation Matrix Deep Dive

Here's how your solution will be scored:


| Criterion | Weight | What They're Judging | Key Risks |
| :-- | :-- | :-- | :-- |
| **Editable PPT** | 30% | Native .pptx with real charts, not screenshots | Embedding images instead of chart objects; poor branding compliance |
| **Adaptability** | 25% | Sector-specific template selection \& metrics | Generic 3-slide structure; wrong metrics for industry |
| **Data Fusion** | 20% | Public + private data coherently blended | Public data missing/irrelevant; contradictions |
| **Anonymization** | 15% | Professional rewriting that obscures identity | Obvious anonymization ("Company ABC"); losing fact accuracy |
| **Citation** | 10% | Valid URLs / file references for all claims | Missing citations; vague attributions |

**Strategy**: Win on Criterion 1 (technically perfect PPT) + Criterion 2 (smart sector logic). These two are worth 55%. Then optimize the others.

***

## 🗺️ End-to-End Processing Flow

```
START: Company Name + Data Pack (Excel/PDFs)
    ↓
[DATA INGESTION]
├─ Parse Excel files → structured JSON
├─ Extract PDF data → tables/text
├─ Scrape company website → raw HTML
└─ Consolidate → unified data structure
    ↓
[SECTOR CLASSIFICATION]
├─ Analyze business description
├─ Extract key metrics
├─ Classify into sector (Manufacturing/D2C/Pharma/Tech/Logistics)
└─ Select appropriate slide template
    ↓
[CONTENT GENERATION - QWEN 3 8B]
├─ Extract headline metrics from private data
├─ Synthesize public web data
├─ Generate anonymized narratives
│  ├─ Slide 1: Business profile
│  ├─ Slide 2: Financial metrics
│  └─ Slide 3: Investment hooks
└─ Prepare chart data for visualization
    ↓
[IMAGE SOURCING]
├─ Generate sector-specific search queries
├─ Fetch from Unsplash/Pexels API
├─ Validate (no logos, good quality)
└─ Store image URLs
    ↓
[CHART GENERATION]
├─ Extract financial metrics
├─ Create native PowerPoint chart objects
│  ├─ Bar: Revenue CAGR + margins
│  ├─ Line: Growth trajectory
│  └─ Pie: Revenue mix / segments
└─ Prepare for embedding
    ↓
[PPT GENERATION - PYTHON-PPTX]
├─ Create 3-slide presentation
├─ Apply Kelp branding (colors, fonts, footer)
├─ Slide 1: Images + anonymized business copy
├─ Slide 2: Images + native charts + KPIs
├─ Slide 3: Investment highlights narrative
└─ Export as editable .pptx
    ↓
[CITATION DOCUMENT]
├─ Compile metadata (claim → source mapping)
├─ Format as Word document
└─ Output .docx
    ↓
END: PPT + Citation Doc (Ready for Submission)
```


***

## 🎯 Sector-Specific Playbooks

### **Manufacturing / Specialty Chemicals**

**Slide 1 - Business Profile \& Infrastructure**

- Visuals: Chemical facilities, production lines, R\&D labs, raw material processing
- Copy: "A leading specialty chemical manufacturer producing [Product Category] for [End-User Industries]"
- Metrics: Product segments (3-5), end-user industries (5-7), manufacturing footprint (number of facilities, states/countries)
- KPIs to highlight: Certifications (GMP+, FSSC 22000, ISO standards), customer count (600+)

**Slide 2 - Financial \& Operational Scale**

- Chart 1 (Bar): Revenue CAGR (target: 15%+), EBITDA margins (target: 20%+)
- Chart 2 (Pie): Revenue mix by segment OR geography (e.g., domestic 55%, exports 45%)
- Copy: Export contribution %, capacity utilization, efficiency metrics
- KPIs: Production capacity (tonnes/year), customer retention %, average order value

**Slide 3 - Investment Highlights**

- Hook 1: "Proprietary product portfolio with high entry barriers"
- Hook 2: "Strategic geographic advantage enabling low logistics costs"
- Hook 3: "Robust financial performance with industry-leading margins"
- Hook 4: (Optional) "Platform for consolidation in fragmented market"

***

### **D2C Consumer Brand**

**Slide 1 - Brand Overview \& Market Presence**

- Visuals: Lifestyle shots of products, packaging (anonymized), digital storefront mockups, customer experience
- Copy: "A direct-to-consumer wellness brand available across [Channels]"
- Metrics: Product portfolio (segments like Men's, Women's, Skincare, Immunity), channel breakdown (Amazon%, Flipkart%, Own website%), certifications (USDA Organic, Non-GMO)
- KPIs: Age of brand, target demographic, SKU count

**Slide 2 - Growth \& Unit Economics**

- Chart 1 (Line): Revenue growth CAGR (target: 65%+) with trajectory over 3-5 years
- Chart 2 (Bar): Unit economics metrics (LTV/CAC ratio > 7x), CAC amount, LTV amount
- Copy: Customer loyalty (repeat rate %), average order value (₹600+), customer acquisition trend
- KPIs: Monthly active customers, net retention rate, payback period for CAC

**Slide 3 - Investment Highlights**

- Hook 1: "Top 3 positioning in key categories on major e-commerce platforms"
- Hook 2: "Profitable operations with industry-leading gross margins (70%+)"
- Hook 3: "Significant market whitespace opportunity (e.g., \$4Bn+ dietary supplements market in India)"
- Hook 4: (Optional) "Strong founder-led team with repeat startup/exit experience"

***

### **Pharma / Life Sciences**

**Slide 1 - Research Pipeline \& Regulatory Status**

- Visuals: Research facilities, clinical labs, product manufacturing
- Copy: "A pharmaceutical company focused on [Therapeutic Area] with [Stage] pipeline"
- Metrics: Pipeline count (pre-clinical/Phase I/II/III/approved), regulatory status, patent portfolio strength
- KPIs: Years to market for lead product, competitive landscape positioning

**Slide 2 - Clinical \& Financial Performance**

- Chart 1 (Bar): Revenue growth from approved products, R\&D efficiency (R\&D spend as % of revenue)
- Chart 2 (Pie): Revenue by therapeutic area / product
- Copy: Clinical trial success rates, time-to-approval historical data, regulatory advantages
- KPIs: Number of approved products, patent exclusivity timeline, revenue per employee

**Slide 3 - Investment Highlights**

- Hook 1: "Patent-protected revenue stream with multi-year exclusivity"
- Hook 2: "Differentiated therapeutic approach addressing unmet medical need"
- Hook 3: "Strong regulatory relationships and approval track record"

***

### **SaaS / Tech**

**Slide 1 - Product \& Market Position**

- Visuals: Product interface, engineering team, technology infrastructure
- Copy: "[Product Name] is a [Category] SaaS platform serving [Customer Segment]"
- Metrics: Addressable market (TAM), customer segment breakdown, feature differentiation
- KPIs: NPS score, time-to-value, product roadmap strength

**Slide 2 - Growth \& Unit Economics**

- Chart 1 (Line): ARR (Annual Recurring Revenue) growth trajectory, retention rate trend
- Chart 2 (Bar): Unit economics (CAC, LTV, Payback period), gross margin %
- Copy: User growth rate, customer acquisition channel mix, product expansion opportunities
- KPIs: Monthly recurring revenue (MRR), customer acquisition cost, lifetime value, magic number

**Slide 3 - Investment Highlights**

- Hook 1: "Strong product-market fit with [NPS/retention metric evidence]"
- Hook 2: "Significant TAM capture opportunity with [X]% penetration potential"
- Hook 3: "Defensible competitive advantage through [network effects / switching costs / proprietary tech]"

***

### **Logistics / Supply Chain**

**Slide 1 - Network \& Infrastructure**

- Visuals: Distribution centers, vehicle fleet, operational hubs, sorting facilities
- Copy: "A logistics solutions provider with presence across [Geographic Coverage]"
- Metrics: Network coverage (cities/states), facility count, vehicle fleet size, geographic footprint
- KPIs: Distribution reach, facility capacity, partnerships with major brands

**Slide 2 - Operational Efficiency \& Scale**

- Chart 1 (Bar): Throughput growth (shipments/year), cost per delivery trend (downward)
- Chart 2 (Line): Utilization rate %, network expansion timeline
- Copy: Turnaround time, delivery reliability (on-time %), cost leadership positioning
- KPIs: Average cost per shipment, delivery radius, automation level

**Slide 3 - Investment Highlights**

- Hook 1: "Capital-efficient business model with improving unit economics at scale"
- Hook 2: "Consolidation play in fragmented logistics market"
- Hook 3: "Essential infrastructure for India's e-commerce and B2B growth"

***

## ⚙️ Technical Stack \& Implementation Plan

### **Recommended Tech Stack**

```
Core Language: Python 3.10+
├─ Web Scraping
│  ├─ BeautifulSoup 4 (parsing HTML)
│  └─ Selenium (JS-heavy sites)
├─ Data Parsing
│  ├─ openpyxl (Excel)
│  ├─ PyPDF2 / pdfplumber (PDF)
│  └─ pandas (data manipulation)
├─ LLM Integration
│  ├─ Ollama / vLLM (Qwen 3 8B inference)
│  └─ Langchain (prompt templates, chains)
├─ Presentation Generation
│  └─ python-pptx (native PPT generation)
├─ Charting
│  ├─ plotly (beautiful charts)
│  └─ openpyxl (to add charts to PPT)
├─ Document Generation
│  ├─ python-docx (Word documents)
│  └─ reportlab (PDF generation)
└─ Image Search
    └─ Unsplash API / Pexels API / Bing Search API

Deployment:
├─ Local LLM: Ollama (easy) or vLLM (optimized)
├─ Memory: Quantized Qwen 3 8B (int8 or fp16)
└─ Orchestration: Python scripts + error handling
```


### **Modular Code Architecture**

```python
projects/
├── data_ingestion/
│   ├── excel_parser.py        # Parse .xlsx files
│   ├── pdf_parser.py          # Extract from PDFs
│   ├── web_scraper.py         # Crawl company websites
│   └── data_consolidator.py   # Merge all sources
├── sector_intelligence/
│   ├── classifier.py          # Detect industry sector
│   └── templates.py           # Sector-specific slide templates
├── content_generation/
│   ├── qwen_interface.py      # Qwen 3 8B inference wrapper
│   ├── anonymizer.py          # Company anonymization
│   └── prompts.py             # LLM prompt templates
├── image_intelligence/
│   ├── query_generator.py     # Sector → image search queries
│   ├── image_fetcher.py       # API integration (Unsplash, etc.)
│   └── image_validator.py     # Logo/quality checks
├── presentation/
│   ├── chart_builder.py       # Create native PPT charts
│   ├── ppt_generator.py       # main PPT assembly
│   └── branding.py            # Kelp brand compliance
├── citation/
│   ├── metadata_tracker.py    # Track data sources
│   └── citation_doc_gen.py    # Word/PDF document creation
├── pipeline.py                # Main orchestrator
└── config.py                  # Global settings
```


***

## 📈 3-Week Development Roadmap

### **Week 1: Foundation \& Data Processing**

- **Goal**: Working data pipeline that can ingest \& structure any company data pack

**Tasks**:

- Set up Qwen 3 8B locally (Ollama or vLLM)
- Build Excel parser (handle multiple sheet formats)
- Build PDF parser (extract tables, text)
- Web scraper for company websites (BeautifulSoup + Selenium)
- Data consolidator to create unified JSON structure
- **Test**: Run on 1-2 dummy companies, verify data extraction accuracy

**Deliverable**: Data ingestion module passes tests on provided company data pack

***

### **Week 2: Intelligence \& Content Generation**

- **Goal**: Sector detection + Qwen-powered content generation working end-to-end

**Tasks**:

- Build sector classifier (rules-based + Qwen-assisted)
- Design sector-specific slide templates
- Integrate Qwen 3 8B for content generation
- Build anonymization engine (prompt-based rewriting)
- Image search integration (Unsplash/Pexels API)
- Start building native PPT generator (basic structure)
- **Test**: Generate rough teasers for all 5 companies (content quality focus)

**Deliverable**: Full content pipeline produces readable 3-slide teasers

***

### **Week 3: Refinement \& Polish**

- **Goal**: Production-ready output with perfect branding, native charts, citations

**Tasks**:

- Polish PPT generation (Kelp branding, layout perfection)
- Implement native chart generation (bar, line, pie)
- Citation document generation (link every claim to source)
- Image quality validation \& fallback strategies
- Extensive testing on all 5 companies
- Performance optimization (inference speed, memory usage)
- Error handling \& logging
- Final cleanup + GitHub repo organization
- **Test**: End-to-end runs on all 5 companies with quality spot-check

**Deliverable**: Production-ready solution ready for submission

***

## 🚀 Implementation Priorities (By Evaluation Weight)

1. **PPT Generation (30%)** → Start early, test extensively
    - Native chart objects are non-negotiable
    - Branding compliance must be pixel-perfect
    - Test with PowerPoint open to verify editability
2. **Adaptability (25%)** → Invest heavily in sector logic
    - Don't build generic template; build adaptive system
    - Test sector classification on diverse companies
    - Verify metrics selection matches sector norms
3. **Data Fusion (20%)** → Quality web scraping + intelligent synthesis
    - Spend time on web scraper reliability
    - Implement smart filtering (not all web content is relevant)
    - Cross-validate public data against private data
4. **Anonymization (15%)** → Prompt engineering + validation
    - Create diverse few-shot examples for Qwen
    - Spot-check all anonymized content manually
    - Balance between obfuscation and readability
5. **Citation (10%)** → Metadata tracking from day 1
    - Every extracted data point → source tag
    - Automate citation document generation
    - Format citations professionally

***

## 🎓 Key Strategic Insights

### **1. Sector Intelligence is Your Differentiator**

Most teams will build a generic "3-slide template" and copy-paste content. You win by **detecting the sector and adapting the entire structure**. Manufacturing needs facility visuals + EBITDA margins. D2C needs lifestyle imagery + LTV/CAC. This adaptive behavior is worth 25% of your score.

### **2. The "Blind" Factor is Surprisingly Hard**

Anonymization isn't just removing names—it's rewriting compelling narratives while obscuring identity. Bad anonymization reads like a legal document. Good anonymization reads like a genuine teaser that happened to omit the company name. Invest in Qwen prompts here.

### **3. Charts Must Be Native, Not Screenshots**

This is a red line. The evaluation explicitly requires "native charts/graphs" in the .pptx. If you embed a PNG bar chart, you'll lose 30% of your score immediately. Native objects (via `python-pptx` chart API) are mandatory.

### **4. Image Quality Matters More Than Perfection**

You don't need 1000+ images; you need 3-4 professional, sector-appropriate images per teaser. Unsplash + Pexels have millions of free, high-quality images. Spend 2-3 minutes per company on image sourcing, not hours.

### **5. Local LLM is a Huge Advantage**

By running Qwen locally, you avoid API costs (staying under ₹100 budget easily) and eliminate dependency on external services. While other teams might rely on GPT-4 API (expensive, risky), you have **offline inference, full privacy, and local control**. This is a competitive edge.

### **6. Citations Are a Scoring Freebie**

Citations are only 10% of the score, but they're the easiest to get perfect. Track data sources from day 1 (every extraction → source tag), then auto-generate citation documents. This is where you can score 9-10/10 reliably.

***

## 🔥 Risk Mitigation \& Contingency Plans

### **Risk 1: Web Scraping Failures (Highest Probability)**

**Why**: Websites change, some have anti-scraping, some require login
**Mitigation**:

- Implement retry logic + exponential backoff
- Fallback to LinkedIn profile if website unavailable
- Cache successfully scraped data
- Have fallback static content for minimal viable teaser


### **Risk 2: Model Hallucination**

**Why**: Qwen might "invent" data not in the provided pack
**Mitigation**:

- Constrain prompts: "Use ONLY the provided data, do not invent facts"
- Implement fact-checking: every generated claim validated against source data
- Use lower temperature (0.3-0.5) for factual tasks, higher (0.7) for rewriting


### **Risk 3: Sector Misclassification**

**Why**: Edge cases (hybrid companies, new industries)
**Mitigation**:

- Multi-classifier ensemble (rules-based + Qwen)
- Confidence score thresholding (if confidence < 0.7, default to generic template)
- Manual override option for ambiguous cases


### **Risk 4: Image Quality Issues**

**Why**: Generic queries might return irrelevant images
**Mitigation**:

- Implement visual validation (check for logos, resolution, relevance)
- Fallback: If good images not found, use simple colored boxes + infographics
- Manual spot-check for 1-2 companies before submission


### **Risk 5: PPT Rendering Issues**

**Why**: Different PowerPoint versions, missing fonts, chart corruption
**Mitigation**:

- Use only standard fonts (Arial, no exotic fonts)
- Test generated PPT on multiple versions of PowerPoint
- Implement fallback chart rendering if native objects fail

***

## 💡 Competitive Advantages

1. **Local LLM Inference** → Zero API dependency, cost efficiency, privacy
2. **Sector-Adaptive Architecture** → Not generic, wins on Adaptability (25%)
3. **Robust Data Ingestion** → Handles multiple formats, hybrid sources
4. **Native PPT Objects** → Editable charts, not screenshots (30% criterion)
5. **Intelligent Anonymization** → Preserves readability while obscuring identity
6. **Professional Branding** → Pixel-perfect Kelp compliance
7. **Complete Citations** → Every metric traceable to source

***

## 📋 Evaluation Readiness Checklist

Before submission, verify:

- [ ] All 5 companies process without errors
- [ ] Each PPT is native, editable .pptx format
- [ ] Charts are PowerPoint objects (not images)
- [ ] Branding is perfect (colors, fonts, footer on every slide)
- [ ] Sector classification accurate for 4-5/5 companies
- [ ] Anonymization removes all identifiers without losing credibility
- [ ] Public web data + private data coherently blended
- [ ] Citation document complete with source references
- [ ] Cost per company < ₹100
- [ ] Processing time reasonable (< 5 min per company)
- [ ] Code is clean, modular, well-documented
- [ ] GitHub repo has comprehensive README

***

This is a **sophisticated problem** that rewards thoughtful architecture, attention to detail, and smart sector logic. Your RTX 4070 + Qwen 3 8B setup is more than sufficient—in fact, the local inference approach is a strategic advantage over teams relying on APIs.

**Good luck! 🚀**
<span style="display:none">[^1]</span>

<div align="center">⁂</div>

[^1]: Problem_Statement.pdf

