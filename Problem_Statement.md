# The Kelp Automated Deal Flow Challenge

**AI M&A Automation Hackathon**  
**Jan 2026**

**Organized by:** Kelp M&A Investment Solutions  
**Prize Pool:** ₹1,60,000  
**Duration:** 3 Weeks

---

## 1. The Problem Statement

In the high-stakes world of Mergers & Acquisitions (M&A), speed and precision are currency. Investment advisors currently spend countless hours manually researching target companies, extracting financial data, and formatting Investment Teasers—brief, 3-slide summaries used to pitch assets to potential buyers anonymously.

Kelp is disrupting this workflow. We are asking you to replace the manual grind with an intelligent AI Agent.

**The Goal:** Build a software pipeline (Python preferred) that accepts a Company Name and a Provided Data Pack as input, and automatically generates a fully formatted, 3-slide Blind Investment Teaser PPT.

---

## 2. The Challenge Workflow (Mandatory Requirements)

Your AI pipeline must autonomously execute the following steps for any given company. Missing any step will result in disqualification.

### 1. Hybrid Data Ingestion
- **Private Data Layer:** The AI must ingest structured datasets provided by Kelp (Excel/PDFs) containing Financials, Credit Reports, and Past Deal Information.
- **Public Data Layer:** For qualitative aspects (Business Model, Products, Market Sentiment), the AI must crawl public sources (Company Website, Blogs).

### 2. Visual Intelligence
The AI must source high-quality, relevant images to visually represent the business.
- **Focus:** Images should depict Products, Manufacturing Plants, or R&D Facilities relevant to the sector.
- **Constraint:** Images must be generic enough to not reveal the specific identity (e.g., no visible logos on factory walls).

### 3. Context-Aware Structuring
The AI must determine the most relevant sections based on the company type (See Section 3 for examples).

### 4. The Blind Factor
All text must be rewritten to anonymize the company while keeping the data accurate.

### 5. Editable PPT Generation
The final output must be a fully editable PowerPoint (.pptx) file containing native charts/graphs where applicable. Strict adherence to the Kelp Branding Guidelines (Colors, Fonts, Footer) must be programmatically enforced.

### 6. Citation Generation
In parallel to the PPT, the code must generate a separate Citation Document (Word/PDF) linking every claim and number to its source.

---

## 3. Sector-Specific Slide Examples

Your AI must detect the sector and populate the slides with the correct "High-Value" information. Below are two examples of how the 3-slide structure should adapt:

### Scenario A: Manufacturing (Specialty Chemicals)
*Reference: Similar to a B2B Ingredient Manufacturer*

**Slide 1: Business Profile & Infrastructure**
- **Key Visuals:** Images of chemical facilities, R&D labs, or raw material processing.
- **Key Sections:** Product Segments (e.g., Lecithin, Phospholipids), End-User Industries (Paints, Pharma, Nutrition), and Manufacturing footprint (e.g., 5 state-of-the-art facilities).

**Slide 2: Financial & Operational Scale**
- **Graphs/Infographics:** A Bar chart showing Revenue Growth (CAGR 15%) and EBITDA Margins (20%).
- **Key Metrics:** Export contribution (e.g., 45% revenue from exports), Certification badges (GMP, FSSC 22000), and Customer count (600+).

**Slide 3: Investment Highlights**
- **The Hook:** Proprietary product portfolio with high entry barriers, Strategic location enabling low logistics costs, and Robust financial performance with industry-leading margins.

### Scenario B: D2C Consumer Brand
*Reference: Similar to a Health/Wellness E-Commerce Brand*

**Slide 1: Brand Overview & Market Presence**
- **Key Visuals:** Lifestyle shots of wellness products, packaging close-ups (anonymized), or digital storefront mockups.
- **Key Sections:** Portfolio Mix (Men's Wellness, Skin Care, Immunity), Channel Presence (Amazon, Flipkart, Own Website), and Certification Logos (USDA Organic, Non-GMO).

**Slide 2: Growth & Unit Economics**
- **Graphs/Infographics:** A chart showing Bottles Sold growth or Revenue Growth (CAGR 65%).
- **Key Metrics:** Customer Loyalty (Repeat Rate 35%), Unit Economics (LTV/CAC 7x), and Average Order Value (₹600).

**Slide 3: Investment Highlights**
- **The Hook:** Ranked Top 3 in key categories on Amazon, Profitable operations with 70% Gross Margins, and Significant whitespace opportunity in the ₹4Bn dietary supplements market.

---

## 4. The Test Suite (5 Companies)

For the final submission, your tool will be tested against 5 Distinct Target Companies across different sectors (e.g., Manufacturing, Tech, Consumer Goods, Pharma, Logistics).

**Logistics:**
- Upon registration, teams will receive a Data Pack for these 5 companies.
- This pack will contain proprietary files (Balance Sheets, Credit Ratings, M&A History) that are not publicly available.
- Your AI must synthesize this Private Data with Public Web Data (blogs/websites) to create the final teaser.

---

## 5. Submission Deliverables

1. **Source Code:** GitHub link or Zip file containing your scripts/agents.
2. **The Output PPTs:** A folder containing the 5 generated .pptx teasers (one for each test company).
3. **The Citation Docs:** 5 separate documents listing sources for the data used.
4. Participants are free to use any free or paid online APIs to achieve results, but the cost incurred must not exceed ₹100 per presentation (approximately 5 slides). Better quality achieved with a lower budget will be rewarded.

---

## 6. Evaluation Matrix

| Criteria | Weightage | Description |
|----------|-----------|-------------|
| Editable PPT Generation | 30% | Is the output a high-quality, editable PowerPoint with native charts/graphs? No static screenshots of text. |
| Adaptability & Sector Logic | 25% | Does the AI choose the right metrics and structure for the specific industry? |
| Data Fusion Capability | 20% | How well does the AI combine the Provided Private Data with Scraped Public Data? |
| Anonymization & Writing | 15% | How smart is the AI at anonymizing and creating content perfect for presentations? |
| Citation Integrity | 10% | Are the claims backed by valid URLs or File references? |

---

## Attachment A: Kelp Branding Guidelines

### 1. Brand Identity
- **Logo:** The Kelp logo (use a text placeholder) must appear on the Top of every slide.
- **Footer:** Every slide must contain the text "Strictly Private & Confidential | Prepared by Kelp M&A Team" in the bottom center (Size 9pt).

### 2. Color Palette
- **Primary (Covers/Overlays):** Dark Indigo/Violet with geometric overlays.
- **Secondary Accents:** Pink-to-Orange Gradient (Brand) and Cyan Blue (Icons).
- **Background:** Clean White for content slides.
- **Text:** White on covers, Dark Grey on body.

### 3. Typography
- **Headings:** Arial Bold (Size 20-24).
- **Body Text:** Arial Regular/Aptos (Size 10-12).

### 4. Layout Principles
- **Information Density:** Avoid large walls of text. Use bullet points and split the slide into 3-4 distinct quadrants or columns.
- **Imagery:** Images must be clean, rectangular, and Full Bleed (stretching to the edge where appropriate), or contained in neat grids.
- **Anonymity:** NO logos of the target company. NO Mention of the specific Company Name in the slides.
