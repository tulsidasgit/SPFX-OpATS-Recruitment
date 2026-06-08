# OpATS — Operational Applicant Tracking System

### A complete walkthrough of the application: what it is, why it exists, and how every piece fits together

---

## 1. Introduction

### 1.1 What is OpATS?

**OpATS** (Operational Applicant Tracking System) is a recruitment management solution built directly on top of **Microsoft 365** — specifically as a **SharePoint Framework (SPFx) web part** that an HR team can drop onto any SharePoint page. It turns SharePoint into a lightweight, no-extra-licensing-cost Applicant Tracking System (ATS), while still leaning on the Microsoft tools recruiters and hiring managers already use every day: **SharePoint Lists, Outlook/Exchange (via Microsoft Graph), Microsoft Forms, and Power Automate** — plus **Anthropic's Claude AI** for intelligent resume screening and job-description drafting.

### 1.2 The problem it solves

Recruitment inside small-to-mid-size organisations is often run through a patchwork of spreadsheets, email threads, and manual resume reading — which leads to:

- Job requisitions that live in someone's inbox instead of a shared, trackable place
- No consistent way to compare candidates against a job's required skills
- Manual, time-consuming first-pass resume screening
- Interview scheduling and feedback collection done over ad-hoc emails, easy to lose track of
- No single "source of truth" for where every candidate stands in the pipeline
- No easy way to look back at closed roles and report on what happened

OpATS addresses every one of these by giving HR **one web part with four purpose-built tabs** that cover the entire lifecycle of a job opening — from posting it, to receiving and screening applications, to interviewing, to closing it out and reporting on the outcome.

### 1.3 Who uses it and how

- **HR / Recruiters** — the primary users. They post jobs, review and screen candidates, schedule interviews, record feedback, and close out completed roles.
- **Hiring managers / interviewers** — receive automated email invitations with interview details and are expected to submit feedback; they don't need direct access to the web part.
- **Candidates** — never touch the SharePoint web part directly. They either fill out a public **Microsoft Form** (external applicants) or are referred/added directly by HR (internal/referral path).
- **"Allowed Posters"** — a controlled subset of HR users (managed via a SharePoint list) who are permitted to create new job postings — preventing uncontrolled job creation.

### 1.4 Why this architecture (the "why Microsoft 365" pitch)

- **No new infrastructure** — everything rides on the SharePoint tenant the company already pays for. No databases, no servers, no hosting bills.
- **No new licensing** — Microsoft Forms and Power Automate are typically already included in Microsoft 365 business plans.
- **Familiar tools** — HR and hiring managers already live in Outlook, SharePoint, and Teams; OpATS meets them there instead of asking them to learn a brand-new SaaS product.
- **Extensible by design** — because it's all SharePoint Lists + Power Automate + a web part, the organisation's own IT/Power Platform team can extend it without needing the original developers.
- **AI where it actually helps** — instead of bolting AI onto everything, OpATS uses Claude surgically: to draft job descriptions and to produce objective, structured fitment reports on resumes — both tasks that are time-consuming and subjective when done manually.

---

## 2. High-Level Architecture

```
                         ┌─────────────────────────────────────────┐
                         │        SharePoint Online (M365)         │
                         │                                         │
   HR User  ───────────► │   OpATS SPFx Web Part (React + Fluent UI)│
   (browser)             │   ┌─────────┬──────────────┬─────────┐ │
                         │   │ Post Job│ Ongoing      │ Track   │ │
                         │   │         │ Positions    │ Progress│ │
                         │   ├─────────┴──────────────┴─────────┤ │
                         │   │        Completed Jobs            │ │
                         │   └──────────────────────────────────┘ │
                         └───────────────┬─────────────────────────┘
                                          │  PnP.js (CRUD)         │ Microsoft Graph
                                          ▼                        ▼
                         ┌─────────────────────────┐   ┌────────────────────────┐
                         │  SharePoint Lists &      │   │  Outlook / Exchange     │
                         │  Document Library        │   │  (Mail.Send via /me)    │
                         │  - Departments            │   │  HR + interviewer       │
                         │  - JobTitles              │   │  notification emails    │
                         │  - JobOpenings            │   └────────────────────────┘
                         │  - Candidates             │
                         │  - Interviews             │   ┌────────────────────────┐
                         │  - AllowedPosters         │   │  Anthropic Claude API   │
                         │  - Resumes (doc library)  │   │  (direct browser calls) │
                         └─────────────────────────┘   │  - Resume fitment        │
                                          ▲              │    screening            │
                                          │              │  - Job description      │
                                          │              │    generation           │
              ┌───────────────────────────┴──────┐       └────────────────────────┘
              │   Microsoft Forms (public form)   │
              │   one form per job posting        │
              └───────────────┬───────────────────┘
                              │ form response
                              ▼
              ┌───────────────────────────────────┐
              │   Power Automate Flow              │
              │   - Maps form fields → Candidates  │
              │   - Stores resume in doc library   │
              │   - Sets initial status fields     │
              └───────────────────────────────────┘
```

### 2.1 Technology stack

| Layer | Technology |
|---|---|
| UI | SharePoint Framework (SPFx) 1.22.2, React 17, Fluent UI |
| Data access | PnP.js (`SPFI`) against SharePoint Lists & Document Libraries |
| Microsoft 365 integration | Microsoft Graph (`/me`, `Mail.Send`) |
| External application intake | Microsoft Forms + Power Automate |
| AI | Anthropic Claude API (browser-direct calls) |
| Build system | Heft (Microsoft's SPFx build toolchain) |

---

## 3. The Application — Tab by Tab

The web part presents itself as a single page with an orange-branded header ("OpATS — Recruitment Tracker") and **four tabs** navigated via a Fluent UI `Pivot` control. Each tab is a self-contained module covering one phase of the recruiting lifecycle.

---

### 3.1 Tab 1 — "Post Job"

**Purpose:** Let authorised HR staff create a new job opening, optionally have AI draft the job description, and generate the public application link.

**Who can use it:** Access is gated. On load, the component checks the current user's email against the **`AllowedPosters`** SharePoint list (`SpService.isAllowedPoster`). If the user isn't on that list, they see a locked screen and cannot post jobs — this prevents anyone with web part access from spamming job postings.

**What the form captures:**
- **Department** and **Job Title** (cascading dropdowns — selecting a department filters the available job titles, both backed by the `Departments` and `JobTitles` SharePoint lists)
- **Required ("Must Have") Skills** and **Good-to-Have Skills**
- **Job Location**, **Job Type** (On Site / Hybrid / Remote), **Experience required**
- **Application Due Date**
- **LinkedIn URL** (optional, for cross-posting reference)
- **Job Description** — written manually, or generated by AI (see below)

**AI-assisted Job Description generation (`AIService.generateJobDescription`):**
1. HR fills in the structural fields (title, department, skills, experience, etc.)
2. HR clicks to open the JD generation panel
3. Behind the scenes, the app:
   - Creates the job opening record immediately (so a real `jobId` exists)
   - Builds and persists the public **Application Form URL** on that job record
   - Optionally fetches a **JD template** file (`JD template.txt` / `JD_template.txt`) stored in the root of the `Resumes` document library — if present, Claude is instructed to mirror its structure, tone, and section headings
   - Sends a structured prompt to Claude, which returns a complete, professional job description (About the Role, Responsibilities, Requirements, What We Offer, How To Apply — including the live application link)
4. HR reviews/edits the generated text and saves it back to the job record

**On submit:**
- The job opening is written to the `JobOpenings` SharePoint list with `Status = Open`
- HR (the full `HR_EMAILS` distribution list configured in `EmailService`) receives a **"New Job Posted"** notification email with a button linking back into the Recruitment Tracker

**Why it matters:** This is the entry point of the whole pipeline — every candidate, screening, and interview downstream traces back to a job created here. The AI drafting feature alone can save HR significant time per posting, and the access gate keeps job creation controlled and auditable.

---

### 3.2 Tab 2 — "Ongoing Positions"

**Purpose:** Give HR a working view of currently active job openings (`Open` or `In Progress`), and provide **two internal application paths** that don't require the candidate to use the public Microsoft Form — useful for referrals and walk-in/direct candidates.

**What HR sees:** A list of active job postings with key details (title, department, location, due date, status).

**Two side-panel actions per job:**

#### a) "Refer Candidate" panel
Used when an employee refers someone for a role. HR fills in:
- Candidate details (name, email, phone, CTC, notice period, etc.)
- **Referrer information** (referrer name, email, employee ID, designation) — these are stored against the candidate record so the organisation can track and potentially reward referrals
- The candidate's resume — uploaded straight from the browser to the `Resumes` SharePoint document library, automatically organised into `Resumes/<Department>/<JobTitle>/Referred/`

#### b) "Apply Directly" panel
Used for walk-in or directly sourced candidates (no referrer). HR fills in the same candidate details minus referrer info, and the resume is stored under `Resumes/<Department>/<JobTitle>/Direct/`.

**What happens after submission (both paths):**
1. `SpService.uploadResume()` ensures the correct folder structure exists (`ensureResumeFolder`) and uploads the file
2. A new record is created in the `Candidates` list with:
   - `FitmentScore = 0`, `ApplicationStatus = Received`, `Source = Referral` or `Direct`
   - All the captured candidate (and referrer, if applicable) details
3. **HR is notified immediately** via `EmailService.notifyHRNewApplication()` — an email summarising the new application and linking back to the tracker

**Duplicate protection:** Before creating a candidate record, the app checks `candidateExistsForJob()` — preventing the same person being logged twice against the same job opening.

**Why it matters:** Not every good candidate comes through a public job posting — referrals and direct walk-ins are often higher-quality and faster to close. This tab makes sure those candidates get into the *same* tracked pipeline as everyone else, with the same visibility and AI screening available to them later.

---

### 3.3 Tab 3 — "Track Progress"

**Purpose:** This is the operational heart of the application — where HR manages every candidate from "received" through to a final hiring decision, for every active job opening.

**Layout:** A list of active job cards (status `Open` / `In Progress`). Clicking a card expands it to reveal its candidates, lazily loaded and grouped into **pipeline stages**:

| Stage | Meaning |
|---|---|
| Pending Screening | Candidate received, AI screening not yet run |
| Screened | AI fitment score has been generated |
| Shortlisted — Round 1 | Candidate has an interview scheduled/completed for round 1 |
| Shortlisted — Round 2 | Progressed to round 2 |
| HR Discussion | Completed round 3 but recommendation isn't yet "Recommended" |
| Final Discussion | Completed round 3 *and* AI recommendation is "Recommended" |
| Rejected | Candidate marked as rejected (can be restored) |

The stage of each candidate is **derived automatically** (`deriveCategory`) from their interview history and fitment data — HR never has to manually move people between buckets; the system reflects reality based on actions taken (screening run, interviews scheduled, feedback submitted, rejection toggled).

**Each candidate is shown as a card with:**
- Name, email, phone
- A colour-coded **fitment score bar** (green ≥ 75%, orange ≥ 50%, red below)
- An **experience-match chip** (`meets` / `exceeds` / `below`)
- **Matching skills** and **missing skills**, shown as chip lists
- The AI-generated **summary** of the candidate's fit
- A **recommendation badge** (Recommended / Maybe / Not Recommended)
- An editable **HR Feedback** notes field (saved independently of everything else)
- A link to **View Resume**
- The candidate's **interview history**, with round number, interviewer, scheduled date/time, and feedback status

**Actions available on each candidate card:**

#### a) "Screen Resume" — AI-powered fitment analysis (`AIService.screenResume`)
This is the headline AI feature. When HR clicks **Screen Resume**:
1. The app fetches the resume content (`SpService.fetchResumeText`)
   - **Plain-text/HTML resumes** are decoded directly
   - **PDF resumes** are sent to Claude as a *native PDF document* (base64-encoded `document` content block) — Claude reads the PDF directly, exactly the way uploading a file to claude.ai works, with no fragile client-side text extraction
2. Claude analyses the resume against the job's required skills, experience, and other requirements, and returns a strict JSON **fitment report** containing: candidate name, contact details, matching skills, missing skills, a 0–100 fitment score, an experience-match verdict, a written summary, and an overall recommendation
3. The app retries automatically (with a corrective follow-up message) if Claude's first response isn't valid JSON
4. The report is written back to the candidate's record in SharePoint, and the card updates immediately to show the new score, chips, summary, and badge
5. **HR is notified by email** (`notifyHRFitmentReady`) that the screening is complete

> **Note:** AI screening is always **HR-triggered** — it never runs automatically on intake. This keeps API usage controlled and gives HR the choice of which candidates are worth the analysis.

#### b) "Schedule Interview" (`InterviewScheduler` panel)
HR selects:
- **Round** (1, 2, or 3)
- **Interviewer** (selected from a people picker resolved via Microsoft Graph)
- **Date and time**

On submission:
- A new record is created in the `Interviews` list (`FeedbackStatus = Pending`)
- The **interviewer receives an email invitation** with the round, candidate, job, date/time, and a link into the tracker (HR is CC'd)
- The candidate's interview history refreshes immediately

#### c) "Reject" / "Restore Candidate"
A one-click toggle that sets `ApplicationStatus = Rejected` (or restores it back to `Received`). Rejected candidates move into their own "Rejected" pipeline stage, and can be restored at any time without losing any of their data (score, history, feedback, etc. are preserved).

#### d) HR Feedback notes
A free-text field HR can use to record their own observations on a candidate, saved independently with its own success/error feedback — useful for capturing context that doesn't fit neatly into the structured AI report.

**Escalation handling:** The `Interviews.FeedbackStatus` field drives an escalation mechanism — when interviewer feedback is overdue, HR receives an escalation email (`notifyHREscalation`) so nothing falls through the cracks. *(The recurring check for this is intended to be implemented as a Power Automate flow — see §6 "What's still to come".)*

**Why it matters:** This tab consolidates everything HR would otherwise track in spreadsheets and email threads — screening results, interview scheduling, feedback status, rejection decisions, and notes — into one live, auto-organising view per job.

---

### 3.4 Tab 4 — "Completed Jobs"

**Purpose:** A read-only retrospective view of every job opening that has been marked **Closed** — for reporting, auditing, and "what happened with this role" conversations.

**How a job gets here:** From the **Track Progress** tab, HR can click **"Close Job"** on any active job card. This opens a confirmation dialog explaining that the job will be removed from Track Progress and moved here (candidates and interview records are fully preserved). On confirmation, `SpService.updateJobOpeningStatus(jobId, 'Closed')` updates the record's status, and it instantly disappears from Track Progress and appears here.

**What HR sees:** The same expandable job-card pattern as Track Progress, but entirely **read-only** — no screening, scheduling, feedback editing, or rejection controls. Each candidate row shows:
- Name, contact details
- An **Outcome badge** — *Selected* (green, when `Recommendation = Recommended` and not rejected), *Rejected* (red), or *Not Selected* (grey) — derived automatically from the stored application status and AI recommendation
- The final fitment score, recommendation, experience-match, matching/missing skills, AI summary, and any HR feedback notes
- The complete interview round history with interviewer, date, and feedback status

**CSV Export:** Each closed job has an **"Export CSV"** button that generates and downloads a spreadsheet of every candidate for that job — including name, contact info, outcome, application status, fitment score, recommendation, experience match, matching/missing skills, AI summary, HR feedback, and a readable summary of their interview history (e.g. *"Round 1: jane@company.com on 12/03/2026 — Submitted; Round 2: …"*). This is generated entirely client-side (no server dependency) and is ready to drop into a report or share with leadership.

**Why it matters:** Recruiting isn't done when a role is filled — being able to look back at exactly how a hiring decision was reached (who applied, how they scored, what interviewers said, why someone was or wasn't selected) is valuable for process improvement, compliance, and reporting on hiring velocity and quality. This tab turns that historical record into something instantly retrievable and exportable.

---

## 4. End-to-End Flows

### 4.1 Flow A — External candidate applies via the public Microsoft Form

```
1. HR posts a job (Tab 1) → ApplicationFormUrl generated & stored on the job
2. HR shares the MS Form link with the outside world (job boards, LinkedIn, company site, etc.)
3. Candidate fills out the public Microsoft Form:
     name · email · phone · CTC · notice period · resume upload · etc.
4. Power Automate flow triggers on the new form response:
     - Maps form fields → a new Candidates list item
     - Stores the resume in Resumes/<Department>/<JobTitle>/Direct/
     - Sets FitmentScore = 0, ApplicationStatus = Received, Source = Direct
     - Optionally bumps the JobOpenings.Status from "Open" to "In Progress"
5. Candidate now appears in Track Progress, ready for HR to review & screen
```

> Note: AI screening is **not** triggered automatically by this flow — it remains an explicit HR action from Track Progress, by design.

### 4.2 Flow B — Internal referral or direct application (via the web part)

```
1. HR opens Ongoing Positions (Tab 2)
2. HR clicks "Refer Candidate" or "Apply Directly" on the relevant job
3. HR fills in candidate details (+ referrer details, if a referral) and uploads the resume
4. App uploads the resume to the correct Resumes/<Dept>/<JobTitle>/{Referred|Direct}/ folder
5. App creates a Candidates record (FitmentScore = 0, ApplicationStatus = Received, Source = Referral/Direct)
6. HR receives an immediate "New Application Received" notification email
7. Candidate now appears in Track Progress, identical to externally-sourced candidates
```

### 4.3 Flow C — AI Resume Screening

```
1. HR opens Track Progress, expands a job, finds a candidate
2. HR clicks "Screen Resume"
3. App fetches the resume:
     - .txt/.html → decoded as plain text
     - .pdf       → base64-encoded and sent to Claude as a native "document" content block
4. Claude reads the resume (and, for PDFs, all pages natively) against the job's
   requirements and returns a strict JSON fitment report
5. App parses the JSON (auto-retrying once with a correction prompt if malformed)
6. Report is saved back to the Candidates list:
     FitmentScore, MatchingSkills, MissingSkills, AISummary, Recommendation, ExperienceMatch
7. Candidate card updates instantly with score bar, chips, summary, and recommendation badge
8. HR receives a "Fitment Report Ready" notification email
```

### 4.4 Flow D — Interview scheduling & feedback escalation

```
1. HR clicks "Schedule Interview" on a candidate card
2. HR selects round (1/2/3), interviewer (people-picker via MS Graph), date & time
3. App creates an Interviews record (FeedbackStatus = Pending)
4. Interviewer receives an email invitation (HR CC'd) with all the details + a tracker link
5. [Pending feedback] → if feedback isn't submitted in time, HR receives an escalation email
6. Once feedback is recorded, FeedbackStatus = Submitted, and the candidate's
   pipeline stage automatically advances (derived live from interview history)
```

### 4.5 Flow E — Closing a job and reviewing the outcome

```
1. HR opens Track Progress, finds a job that has reached its conclusion
2. HR clicks "Close Job" → confirms in the dialog
3. App sets JobOpenings.Status = Closed
     → Job instantly disappears from Track Progress
     → Job instantly appears in Completed Jobs (read-only)
4. HR (or leadership) opens Completed Jobs, expands the job, reviews every
   candidate's final outcome, score, interview history, and feedback
5. HR clicks "Export CSV" to download a full spreadsheet record of the hiring round
```

---

## 5. Behind the Scenes — Services Layer

The web part is built on a small set of focused services (all instantiated per-component, no global state management — just class components with local state):

| Service | Responsibility |
|---|---|
| **`SpService`** | All SharePoint CRUD: job openings, candidates, interviews, departments/job titles, allowed-posters check, resume upload & retrieval, JD template retrieval |
| **`GraphService`** | Current user info (`/me`) and sending mail (`/me/sendMail`) via Microsoft Graph with delegated permissions |
| **`AIService`** | Two Claude-powered capabilities: `screenResume()` (resume fitment analysis, including native PDF document support) and `generateJobDescription()` (AI-drafted JD from form fields + optional template) |
| **`EmailService`** | Five HTML email notification scenarios, all routed through `GraphService` |

### 5.1 Email notifications — when HR (and others) get pinged

| Trigger | Recipient(s) | Purpose |
|---|---|---|
| New job posted | HR distribution list | Confirms a new requisition is live |
| New application received (referral/direct) | HR distribution list | Surfaces internally-sourced candidates immediately |
| AI fitment report ready | HR distribution list | Tells HR a screening just completed and is ready to review |
| Interview scheduled | Interviewer (HR CC'd) | Gives the interviewer everything they need: round, candidate, job, date/time, tracker link |
| Feedback escalation | HR distribution list | Flags overdue interview feedback so it doesn't get lost |

### 5.2 The AI, in plain terms

OpATS uses Anthropic's **Claude** model for two distinct, well-scoped jobs — never as a generic chatbot bolted onto the UI:

1. **Resume screening** — given a job's requirements and a candidate's resume (plain text *or* a native PDF document), Claude returns a consistent, structured JSON verdict: a 0–100 fitment score, matched/missing skills, an experience-match classification, a short written summary, and an overall recommendation (Recommended / Maybe / Not Recommended). The prompt is explicitly engineered to:
   - Always return complete, valid JSON (with sensible fallback values for missing fields) — minimizing parsing failures
   - Gracefully handle unreadable resumes (e.g. scanned image-only PDFs) by returning a clear, structured "could not be read" report instead of failing outright
   - Stick to a fixed set of allowed values for fields like `recommendation` and `experienceMatch`, so the UI can render them consistently every time

2. **Job description generation** — given the structural details of a role (and, optionally, a company JD template file), Claude drafts a complete, professional, ready-to-publish job description — including a "How to Apply" section that embeds the live application link.

---

## 6. Access Control & Data Model Summary

### 6.1 Who can do what

- **Posting jobs** is gated by the `AllowedPosters` SharePoint list — only listed email addresses can create job openings; everyone else sees a locked view.
- **Everything else** (referring/applying for candidates, screening, scheduling, closing jobs) is available to anyone with access to the web part — i.e., the broader HR team.
- **Candidates and interviewers never access the web part** — they interact purely through the Microsoft Form (candidates) or email invitations (interviewers).

### 6.2 SharePoint Lists & Library (the data backbone)

| List / Library | Purpose |
|---|---|
| `Departments` | Organisational departments, used to scope job titles and folder structure |
| `JobTitles` | Job titles, linked to departments (cascading dropdown source) |
| `JobOpenings` | The job requisitions themselves — title, department, skills, status, due date, JD, application URL, etc. |
| `Candidates` | Every applicant — contact info, resume link, AI fitment fields, HR feedback, recommendation, referral details, application status |
| `Interviews` | Every scheduled interview round — interviewer, date, feedback status, feedback text, HR notes |
| `AllowedPosters` | Controls who is permitted to post new jobs |
| `Resumes` (document library) | All resume files, auto-organised as `Resumes/<Department>/<JobTitle>/{Direct, Referred}/` |

---

## 7. What Makes This Worth Presenting

- **It replaces a fragmented, manual process with one connected pipeline** — from "a role needs filling" all the way to "here's a CSV of exactly how we filled it"
- **It uses AI exactly where it adds the most value** — drafting JDs and producing consistent, explainable fitment scores — without making the system feel like "just an AI demo"
- **It costs nothing extra to run** — no new servers, databases, or per-seat SaaS licensing; it's built entirely from tools already inside a Microsoft 365 tenant
- **It's transparent and auditable** — every status change, score, email, and interview outcome is stored as a normal SharePoint list item, queryable and reportable by IT or leadership at any time
- **It scales with the organisation's existing Power Platform skills** — anyone who can build a Power Automate flow or edit a SharePoint list can extend it further

---

## 8. What's Still To Come

A few pieces of the original project plan remain open (tracked in `CLAUDE.md`):

- **Graph API permission requests** — formally declaring `User.Read`, `Mail.Send`, `Sites.ReadWrite.All`, and `Files.ReadWrite.All` in the solution package so admin consent can be granted cleanly during deployment
- **Power Automate flow definitions** — codifying the interview-feedback escalation flow (recurring check, e.g. every 3 hours) and the form-response-to-candidate flow as exportable, documented flow definitions
- **Full documentation set** — developer guide, architecture deep-dive, user guide, and a deployment checklist, so the system can be handed off and maintained independently
- **Polished `README`** — a proper project landing page with setup instructions and an architecture diagram

None of these affect what the application *does* today — they're about making it easier to **deploy, hand off, and maintain** going forward.
