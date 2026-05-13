# ISG Digital Onboarding Portal

**Integrated Sycamore Group LLC** — HCS Provider · Assisted Living & Residential Care · Houston, TX

A fully self-contained digital employment application and regulatory onboarding package built as two standalone HTML files. No server, no framework, no cost — open in any browser, fill out, and save as PDF.

Built as a practical alternative to an expensive website or fillable Word document. The entire solution is two HTML files that can be emailed, shared via link, or hosted on any static file host (GitHub Pages, Google Drive, Dropbox, etc.).

---

## 📁 Project Files

```
isg-onboarding/
├── onboarding_portal.html      ← Main employment application (applicant fills this)
├── letter_of_reference.html    ← Standalone referee form (sent to each referee)
└── README.md
```

---

## 📋 What the Portal Covers

The main application (`onboarding_portal.html`) contains **13 sections** in a single scrollable page:

| # | Section | Key Fields |
|---|---------|-----------|
| 1 | Position & Application | Title, pay rate, employment type, start date, how-heard checkboxes |
| 2 | Personal Information | Name, SSN, current & previous address, emergency contact, eligibility Y/N, DL details, background disclosure + conviction table |
| 3 | Availability | Days (Mon–Sun + Holidays), shift-time checkboxes (6 shift blocks) |
| 4 | Education History | Free-input table with example row, add-row button |
| 5 | Licenses & Certifications | License table (expandable), CPR/ACLS/Malpractice, HCS competency checkboxes |
| 6 | Employment History | 3 employer blocks — duties, pay, supervisor, contact consent |
| 7 | Character References | 3 referee profiles with name, relationship, contact |
| 8 | Health & Compliance (HHSC) | TB, COVID status, background check consent + signed authorization |
| 9 | Certification & Signature | 4 statutory statements + draw-to-sign canvas |
| 10 | HCS Competency Evaluation | 9 essay questions + reviewer block |
| 11 | ANE Acknowledgement | SB 2170 full legal text + 3-column signature table |
| 12 | Complaint Procedure | 5-step process, DADS contact directory, dual manager signatures |
| 13 | Letters of Reference | Referee contact log + instructions to send `letter_of_reference.html` |
| 14 | IRS Form W-4 (2026) | All 5 steps complete + Employer Use Only block |

---

## ✅ Features

### Form Functionality
- **Fully fillable** — every field has a label and placeholder guidance text
- **Draw-to-sign** — canvas signature pads on all required signature lines (mouse, trackpad, and touch)
- **Yes/No radio buttons** — for all eligibility, consent, and background questions
- **Checkbox groups** — availability days/shifts, HCS competencies, how-heard, W-4 filing status
- **Auto-expanding tables** — Conviction, License, and Education tables each gain a new row automatically when the last row is fully filled; manual `+ Add Row` buttons also available
- **Education example row** — greyed-out italic example disappears on first click, leaving clean blank rows

### Technical
- **Zero dependencies** — pure HTML, CSS, and vanilla JS; no npm, no CDN, no build step
- **Single file per form** — everything embedded; just open and use
- **Responsive grid** — 12-column CSS grid collapses cleanly on mobile
- **Print stylesheet** — removes UI chrome, preserves field borders for clean PDF output
- **Touch support** — signature canvases work on phones and tablets

### Compliance
- Texas Health & Safety Code Ch. 250 (ANE mandatory reporting)
- Texas SB 2170 (full acknowledgement text)
- Texas HHSC HCS Provider Standards
- HIPAA consumer confidentiality
- IRS Form W-4 (2026, all 5 steps per Publication 15-T)
- DFPS background check authorization

---

## 🔗 Letter of Reference Flow

The letter of reference is a **separate standalone file** (`letter_of_reference.html`):

```
Applicant fills onboarding_portal.html
    ↓
Section 13 — applicant notes referee names and contact info
    ↓
Applicant sends letter_of_reference.html link/file to each referee
    ↓
Each referee opens the file independently, fills in their details,
writes their letter, draws their signature, saves as PDF
    ↓
Referee emails completed PDF back to ISG
```

One file, shared three times — no duplication, no confusion.

---

## 🚀 Usage

### Option A — Open Directly (no server needed)
```bash
# macOS
open onboarding_portal.html

# Windows
start onboarding_portal.html

# Linux
xdg-open onboarding_portal.html
```

### Option B — Host on GitHub Pages
1. Fork this repo
2. Go to **Settings → Pages → Source → main branch**
3. Your portal is live at `https://yourusername.github.io/isg-onboarding/onboarding_portal.html`
4. Share that URL with applicants

### Option C — Host on any static file host
Upload both HTML files to Google Drive, Dropbox, OneDrive, Netlify, or any web host. No backend required.

### Saving a Completed Form as PDF
1. Fill out all fields
2. Draw signatures in all gold-bordered signature areas
3. Click **"Download Completed Application (PDF)"**
4. In the print dialog → set **Destination: Save as PDF**
5. Paper: **Letter**, Margins: **Minimum** or **None**
6. Save

---

## 🎨 Customisation

All brand colours are CSS variables at the top of the `<style>` block in each file:

```css
:root {
    --navy:     #1B3A5C;   /* Section headers, borders */
    --gold:     #C8973A;   /* Accent colour, signature borders */
    --bg:       #F0F3F7;   /* Page background */
    --label-bg: #E8ECF0;   /* Field label strip */
    --input-bg: #FAFAF8;   /* Input field background */
    --success:  #1e6b2e;   /* Submit / download button */
}
```

To update company details, find and replace `Integrated Sycamore Group LLC` and the address/phone values.

---

## ⚖️ Regulatory Reference

| Regulation | Where Covered |
|-----------|--------------|
| Texas Health & Safety Code Ch. 250 | Section 2 (background disclosure), Section 8 (signed authorization) |
| Texas SB 2170 / HRC Ch. 48 | Section 11 — full statutory ANE text, 3-party signatures |
| HHSC HCS Standards (40 TAC Ch. 9) | Sections 4, 5, 10 — education, licensure, competency evaluation |
| HIPAA (45 CFR Parts 160 & 164) | Section 10 — Evaluation Q5 consumer confidentiality |
| IRS Publication 15-T (2026) | Section 14 — W-4 all 5 steps, 2026 credit amounts |
| DFPS Background Check | Section 8 — signed authorization paragraph |

> **Disclaimer:** This form is an administrative intake tool only. It does not constitute legal advice. Verify all regulatory requirements against current HHSC guidelines before use.

---

## 💡 Why HTML Instead of Word or a Website

| Approach | Cost | Works offline | No install | Signable | Shareable by email |
|---------|------|--------------|------------|----------|-------------------|
| This HTML portal | Free | ✅ | ✅ | ✅ (canvas) | ✅ |
| Fillable Word (.docx) | Free | ✅ | ❌ (needs Word/LibreOffice) | ❌ | ✅ |
| Google Forms | Free | ❌ | ✅ | ❌ | ✅ |
| Custom website | $$$+ | ❌ | ✅ | Varies | ✅ |
| PDF form | Free | ✅ | ❌ (needs Acrobat) | Varies | ✅ |

The HTML approach hits every requirement: free, no software install, works offline, supports draw-to-sign, and can be emailed as a single file or shared as a URL.

---

## 📞 Contact

**Integrated Sycamore Group LLC**
8315 Sierra Hill Ct, Houston TX 77083
Tel: 832-335-4098 / 713-835-2122
