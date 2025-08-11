# PowerPoint Consistency Checker

This tool uses AI to analyze PowerPoint presentations for factual and logical inconsistencies across slides, such as conflicting numerical data, contradictory claims, and timeline mismatches.

## Features & Capabilities

**Multi-Slide Analysis:** Detects inconsistencies across all slides in a presentation

**Comprehensive Checks:**
- Conflicting numerical data (values, percentages)
- Contradictory textual claims
- Timeline mismatches
- Internal consistency issues (totals vs breakdowns)
- Comparative claim validation

**Multi-Modal Extraction:**
- Text from shapes and tables
- OCR for text in images
- Slide title preservation

**Prioritized Output:**
- **Critical:** Major numerical conflicts or core claim contradictions
- **High:** Significant inconsistencies affecting key messages
- **Medium:** Minor mismatches requiring clarification
- **Low:** Presentation inconsistencies without material impact

## Installation

### Prerequisites
- Python 3.9+
- Google Gemini API key ([Get API key](https://aistudio.google.com/app/apikey))

### Setup

\`\`\`bash
# Clone repository
git clone https://github.com/sandeep14k/nougat.git
cd nougat

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate    # Windows

# Install dependencies
pip install -r requirements.txt

# Install Tesseract OCR (Mac)
brew install tesseract

# Install Tesseract OCR (Windows)
# Download installer: https://github.com/UB-Mannheim/tesseract/wiki
\`\`\`

## Configuration

1. Set your Gemini API key as environment variable:

\`\`\`bash
# Linux/Mac
export GOOGLE_API_KEY="your_api_key_here"

# Windows (Powershell)
$env:GOOGLE_API_KEY="your_api_key_here"
\`\`\`

2. (Optional) For better OCR accuracy, install Tesseract language data:

\`\`\`bash
# English only
sudo apt-get install tesseract-ocr-eng  # Debian/Ubuntu

# All languages
brew install tesseract-lang  # Mac
\`\`\`

## Usage

\`\`\`bash
python ppt_analyzer.py [INPUT.pptx] [-o OUTPUT.json] [-v]

# Basic analysis
python ppt_analyzer.py presentation.pptx

# Save results to file
python ppt_analyzer.py presentation.pptx -o results.json

# Verbose mode (debug logging)
python ppt_analyzer.py presentation.pptx -v
\`\`\`

## Sample Output (results.json)

\`\`\`json
{
  "statistics": {
    "total_slides": 7,
    "issues_found": 5,
    "analysis_time": "2025-08-12 15:22:18"
  },
  "inconsistencies": [
    {
      "type": "Contradictory core claims",
      "slides": [1, 2],
      "description": "Critical conflict in speed improvement claims: 2x faster vs 3x faster",
      "evidence": [
        "Slide 1 TITLE: Noogat helps consultants make decks 2x faster using AI",
        "Slide 2: 3x faster deck creation speed"
      ],
      "severity": "Critical"
    }
  ]
}
\`\`\`

## Functionality

**Extraction Phase:**
- Processes PPTX file using python-pptx
- Extracts text from shapes, tables, and titles
- Performs OCR on images using Tesseract
- Structures content with slide context markers

**Analysis Phase:**
- Sends structured content to Gemini 1.5 Flash
- Uses specialized prompt for inconsistency detection
- Handles API rate limits with exponential backoff
- Processes response to extract JSON-formatted results

**Output Phase:**
- Sorts findings by severity
- Generates summary statistics
- Outputs machine-readable JSON

