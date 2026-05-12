# 🤖 HIREVIA — AI-Powered Resume Screening System

> **Pay-per-resume AI shortlisting for recruitment agencies in US, UK & Australia.**
> No subscriptions. No lock-in. Just results.

---

## 🌟 What Is HIREVIA?

HIREVIA is an AI-powered resume screening system that evaluates, scores, and ranks candidates against a job description — in minutes, not hours.

Built for recruitment agencies who are tired of:
- Spending hours manually reading resumes
- Paying $10,000+/year for enterprise ATS tools
- Getting biased, inconsistent shortlists

**With HIREVIA — send us your JD + resumes. Get back a ranked Excel report and sorted PDFs. Pay only $0.30 per resume.**

---

## ⚡ Why HIREVIA Over Traditional ATS?

| Feature | Greenhouse | Lever | Workable | **HIREVIA** |
|---|---|---|---|---|
| Price | $25,000/yr | $15,000/yr | $375/mo | **$0.30/resume** |
| Contract | Annual | Annual | Monthly | **None** |
| AI Scoring | ❌ Basic | ❌ Basic | ❌ Basic | ✅ Deep AI |
| Bias Removal | ❌ No | ❌ No | ❌ No | ✅ Built-in |
| Transparent Scoring | ❌ Black box | ❌ Black box | ❌ Black box | ✅ Full breakdown |
| Setup Time | Weeks | Weeks | Days | **Zero** |
| Pay Per Use | ❌ | ❌ | ❌ | ✅ |

---

## 🎯 Core Features

### ✅ ATS-Style Hard Filter
Automatically flags candidates who don't meet:
- Minimum years of experience
- Required skills
- Location requirements
- Work authorization requirements

> ⚠️ Candidates are **flagged, never rejected** — recruiter always makes final call.

### 🤖 AI Evaluation (Groq — llama-3.3-70b)
Each resume is independently scored across 5 dimensions:

| Dimension | Weight | What It Measures |
|---|---|---|
| Skill Match | **40%** | JD-required skills found in resume |
| Experience Relevance | **20%** | How closely past roles match JD |
| Project Relevance | **20%** | Alignment of projects with role |
| Domain Fit | **10%** | Industry/sector alignment |
| Career Progression | **10%** | Growth and increasing responsibility |

### 🛡️ Bias Removal (Built-In)
Before AI evaluation, the system automatically removes:
- ❌ Candidate name
- ❌ Email address
- ❌ Phone number
- ❌ Gender pronouns
- ❌ Age / Date of birth
- ❌ Physical address
- ❌ Nationality / Citizenship
- ❌ Photo references

**Evaluation is 100% skill and experience based.**

### 📊 Two Client-Ready Outputs

**Output 1 — ranked_candidates.xlsx**
- Professional Excel table
- Sorted highest → lowest score
- Colour coded (green = top 3, yellow = top 10)
- Clickable resume links
- Auto-filters for easy sorting
- Summary dashboard sheet

**Output 2 — ranked_resumes.zip**
- All resumes renamed by rank
- Example: `Rank001_91pct_Anna_Brown.pdf`
- Best candidate always first
- ZIP ready to send to client

---

## 🚀 Quick Start

### 1. Clone the repository
```bash
git clone https://github.com/vishakhachaudhari24/HIREVIA.git
cd HIREVIA
```

### 2. Create virtual environment
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Mac/Linux
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Set up API key
```bash
copy .env.example .env
```
Open `.env` and add your Groq API key:
```
GROQ_API_KEY=gsk_xxxxxxxxxxxxxxxxxxxxxxxx
```
Get your free key at **console.groq.com**

### 5. Run
```bash
python main.py jd.txt resumes/
```

---

## 📁 Project Structure

```
HIREVIA/
├── main.py               ← Entry point — full pipeline
├── scorer.py             ← Weighted scoring formula (pure Python)
├── output_generator.py   ← Excel + ZIP generation
├── requirements.txt      ← Dependencies
├── .env.example          ← API key template
├── sample_jd.txt         ← Example job description
└── README.md             ← This file
```

---

## 💻 Usage

```bash
python main.py <job_description.txt> <resumes_folder_or_zip>
```

**Examples:**
```bash
# Using a folder
python main.py jd.txt ./resumes/

# Using a ZIP file
python main.py jd.txt resumes.zip
```

---

## 📊 Output Preview

### Excel Report (ranked_candidates.xlsx)

| Rank | Candidate | Score | Criteria Met | Key Skills | Gaps | AI Summary |
|---|---|---|---|---|---|---|
| Rank001 | Anna Brown | 91% | Yes | Python, AWS, Docker | Redis | Strong backend engineer... |
| Rank002 | John Smith | 84% | Yes | Python, REST API | Fintech exp | Solid developer... |
| Rank003 | Mike Patel | 45% | No | Python | Docker, PostgreSQL | Different domain... |

### Ranked PDFs (ranked_resumes.zip)
```
Rank001_91pct_Anna_Brown.pdf     ← Best match
Rank002_84pct_John_Smith.pdf
Rank003_45pct_Mike_Patel.pdf
```

---

## 🧮 Scoring Formula

```
Final Score = (0.40 × Skill Match)
            + (0.20 × Experience Relevance)
            + (0.20 × Project Relevance)
            + (0.10 × Domain Fit)
            + (0.10 × Career Progression)
```

Calculated in `scorer.py` using pure Python — **no AI involved in scoring.**

---

## ⚙️ Configuration

In `main.py` you can configure:

```python
MAX_RESUME_CHARS = 6000   # Max chars sent to AI per resume
TOP_N = None              # Set to 20 for top 20 only in Excel
```

For large batches (500+ resumes), increase delay:
```python
time.sleep(1.0)   # Change from 0.5 to 1.0
```

---

## 💰 Pricing Model

| Volume | Price | Your Revenue |
|---|---|---|
| 100 resumes | $0.30/resume | $30 |
| 500 resumes | $0.30/resume | $150 |
| 1,000 resumes | $0.30/resume | $300 |
| 5,000 resumes | $0.30/resume | $1,500 |

**API cost at 1,000 resumes: ~$0.50**
**Your margin: ~98%**

---

## 🛠️ Tech Stack

| Component | Technology |
|---|---|
| Language | Python 3.10+ |
| AI Model | Groq — llama-3.3-70b-versatile |
| PDF Extraction | pdfplumber |
| Excel Generation | openpyxl + pandas |
| Environment | python-dotenv |

---

## ⚠️ Troubleshooting

| Error | Fix |
|---|---|
| `GROQ_API_KEY not set` | Check `.env` file exists with correct key |
| `No PDF files found` | Make sure resumes are `.pdf` format |
| `Empty PDF warning` | PDF is scanned/image-based — needs OCR |
| `JSON parse error` | Rare — retry the script |
| `venv not activating` | Run `Set-ExecutionPolicy RemoteSigned` |

---

## 🗺️ Roadmap

- [x] PDF text extraction
- [x] PII / bias removal
- [x] ATS hard filter
- [x] AI evaluation (Groq)
- [x] Weighted scoring formula
- [x] Excel output with formatting
- [x] Ranked PDF + ZIP output
- [ ] Web application (self-serve)
- [ ] Stripe payment integration
- [ ] Client dashboard
- [ ] Multi-language resume support
- [ ] OCR for scanned PDFs
- [ ] API endpoint for integrations

---

## 📬 Contact

Built by **Vishakha Chaudhari**

For partnerships or enterprise inquiries:
- GitHub: [@vishakhachaudhari24](https://github.com/vishakhachaudhari24)

---

## ⚖️ License

This project is proprietary and confidential.
Unauthorized copying, distribution, or use is strictly prohibited.

---

<p align="center">
  <b>HIREVIA — Screen smarter. Hire faster.</b>
</p>
