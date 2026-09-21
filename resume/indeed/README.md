# Indeed resume — Anthony Ty Marler

Upload this file to Indeed:

**[`Anthony_Ty_Marler_Indeed_Resume.docx`](Anthony_Ty_Marler_Indeed_Resume.docx)**

Indeed also accepts the PDF if a posting asks for one:

**[`Anthony_Ty_Marler_Indeed_Resume.pdf`](Anthony_Ty_Marler_Indeed_Resume.pdf)**

## Why this file

Indeed converts the upload to a PDF and also parses it into your profile. This version follows Indeed’s ATS guidance:

- `.docx` first (safest parse), text-based PDF as backup
- Single column, no tables, images, headers, or footers
- Standard headings: Summary, Skills, Experience, Education, Patents
- Dates on their own line (`March 2025 - Present`) so the parser does not attach them to the wrong job
- Reverse-chronological work history from Independent (2025) through U.S. Army (1984–1987)

Content matches the latest resume on `cursor/resume-generation-28e1` plus later LinkedIn copy (Fractional CTO headline, four Independent AI applications, Army February 1984–February 1987).

## How to upload

1. Sign in at [indeed.com](https://www.indeed.com/)
2. Open **Profile** / **Resume**
3. Choose **Upload resume**
4. Select `Anthony_Ty_Marler_Indeed_Resume.docx`
5. Preview the parsed fields (name, phone, email, jobs, dates) before saving

File size is well under Indeed’s typical 2–5 MB apply cap.

## Regenerate

```bash
python3 resume/indeed/generate_indeed_resume.py
```

PDF (optional, needs Chrome):

```bash
google-chrome --headless --disable-gpu --no-pdf-header-footer \
  --print-to-pdf=resume/indeed/Anthony_Ty_Marler_Indeed_Resume.pdf \
  resume/indeed/Anthony_Ty_Marler_Indeed_Resume.html
```
