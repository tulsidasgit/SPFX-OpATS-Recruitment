import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/folders';
import '@pnp/sp/files';
import {
  IDepartment,
  IJobTitle,
  IJobOpening,
  ICandidate,
  IInterview,
  IFitmentReport,
} from './models';

export class SpService {
  private _sp: SPFI;
  private _webUrl: string | null = null;
  private _resumesLibraryUrl: string | null = null;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  // ── Departments ────────────────────────────────────────────────────────────

  public async getDepartments(): Promise<IDepartment[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('Departments')
        .items.select('Id', 'Title', 'IsActive')
        .filter('IsActive eq 1')
        .orderBy('Title')();

      return items.map((i: { Id: number; Title: string; IsActive: boolean }) => ({
        id: i.Id,
        title: i.Title,
        isActive: i.IsActive,
      }));
    } catch (err) {
      throw new Error(`Failed to load departments: ${(err as Error).message}`);
    }
  }

  // ── JobTitles ──────────────────────────────────────────────────────────────

  public async getJobTitlesByDepartment(departmentId: number): Promise<IJobTitle[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('JobTitles')
        .items.select('Id', 'Title', 'DepartmentId')
        .filter(`DepartmentId eq ${departmentId}`)
        .orderBy('Title')();

      return items.map((i: { Id: number; Title: string; DepartmentId: number }) => ({
        id: i.Id,
        title: i.Title,
        departmentId: i.DepartmentId,
      }));
    } catch (err) {
      throw new Error(`Failed to load job titles: ${(err as Error).message}`);
    }
  }

  // ── JobOpenings ────────────────────────────────────────────────────────────

  public async getOpenJobOpenings(): Promise<IJobOpening[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('JobOpenings')
        .items.select('Id', 'Title', 'Department', 'JobTitle', 'RequiredSkills', 'GoodToHaveSkills', 'JobLocation', 'JobType', 'Experience', 'DueDate', 'Status', 'PostedBy', 'LinkedInUrl', 'JobDescription')
        .filter("Status eq 'Open' or Status eq 'In Progress'")
        .orderBy('Created', false)();

      return items.map(this._mapJobOpening);
    } catch (err) {
      throw new Error(`Failed to load job openings: ${(err as Error).message}`);
    }
  }

  public async getClosedJobOpenings(): Promise<IJobOpening[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('JobOpenings')
        .items.select('Id', 'Title', 'Department', 'JobTitle', 'RequiredSkills', 'GoodToHaveSkills', 'JobLocation', 'JobType', 'Experience', 'DueDate', 'Status', 'PostedBy', 'LinkedInUrl', 'JobDescription')
        .filter("Status eq 'Closed'")
        .orderBy('Modified', false)();

      return items.map(this._mapJobOpening);
    } catch (err) {
      throw new Error(`Failed to load closed jobs: ${(err as Error).message}`);
    }
  }

  public async createJobOpening(job: Omit<IJobOpening, 'id'>): Promise<number> {
    try {
      const result = await this._sp.web.lists.getByTitle('JobOpenings').items.add({
        Title: job.title,
        Department: job.department,
        JobTitle: job.jobTitle,
        RequiredSkills: job.requiredSkills,
        GoodToHaveSkills: job.goodToHaveSkills,
        JobLocation: job.jobLocation,
        JobType: job.jobType,
        Experience: job.experience,
        DueDate: job.dueDate,
        Status: job.status,
        PostedBy: job.postedBy,
        LinkedInUrl: job.linkedInUrl ?? '',
      });
      return (result.data as { Id: number }).Id;
    } catch (err) {
      throw new Error(`Failed to create job opening: ${(err as Error).message}`);
    }
  }

  public async updateJobDescription(id: number, jobDescription: string): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle('JobOpenings').items.getById(id).update({
        JobDescription: jobDescription,
      });
    } catch (err) {
      throw new Error(`Failed to save job description: ${(err as Error).message}`);
    }
  }

  public async getJobOpeningById(id: number): Promise<IJobOpening> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('JobOpenings')
        .items.select('Id', 'Title', 'Department', 'JobTitle', 'RequiredSkills', 'GoodToHaveSkills', 'JobLocation', 'JobType', 'Experience', 'DueDate', 'Status', 'PostedBy', 'LinkedInUrl', 'JobDescription')
        .filter(`Id eq ${id}`)
        .top(1)();
      if (items.length === 0) throw new Error(`Job opening #${id} not found.`);
      return this._mapJobOpening(items[0] as Record<string, unknown>);
    } catch (err) {
      throw new Error(`Failed to load job opening: ${(err as Error).message}`);
    }
  }

  public async updateJobOpeningApplicationUrl(id: number, url: string): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle('JobOpenings').items.getById(id).update({
        ApplicationFormUrl: url,
      });
    } catch {
      // Silently skip — ApplicationFormUrl column may not exist on all tenants
    }
  }

  public async updateJobOpeningStatus(id: number, status: IJobOpening['status']): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle('JobOpenings').items.getById(id).update({ Status: status });
    } catch (err) {
      throw new Error(`Failed to update job status: ${(err as Error).message}`);
    }
  }

  // ── AllowedPosters ─────────────────────────────────────────────────────────

  public async isAllowedPoster(email: string): Promise<boolean> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('AllowedPosters')
        .items.select('UserEmail')
        .filter(`UserEmail eq '${email}'`)
        .top(1)();
      return items.length > 0;
    } catch (err) {
      throw new Error(`Failed to check poster permissions: ${(err as Error).message}`);
    }
  }

  // ── Candidates ─────────────────────────────────────────────────────────────

  public async getCandidatesByJobOpening(jobOpeningId: number): Promise<ICandidate[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('Candidates')
        .items.select(
          'Id', 'JobOpeningId', 'CandidateName', 'Email', 'Phone', 'ResumeUrl',
          'FitmentScore', 'MatchingSkills', 'MissingSkills', 'AISummary',
          'HRFeedback', 'Recommendation', 'ExperienceMatch', 'ApplicationStatus'
        )
        .filter(`JobOpeningId eq ${jobOpeningId}`)
        .orderBy('FitmentScore', false)();

      return items.map(this._mapCandidate);
    } catch (err) {
      throw new Error(`Failed to load candidates: ${(err as Error).message}`);
    }
  }

  public async getAllCandidates(): Promise<ICandidate[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('Candidates')
        .items.select(
          'Id', 'JobOpeningId', 'CandidateName', 'Email', 'Phone', 'ResumeUrl',
          'FitmentScore', 'MatchingSkills', 'MissingSkills', 'AISummary',
          'HRFeedback', 'Recommendation', 'ExperienceMatch', 'ApplicationStatus'
        )
        .top(5000)();

      return items.map(this._mapCandidate);
    } catch (err) {
      throw new Error(`Failed to load candidates: ${(err as Error).message}`);
    }
  }

  public async candidateExistsForJob(email: string, jobId: number): Promise<boolean> {
    try {
      const safe = email.replace(/'/g, "''");
      const items = await this._sp.web.lists
        .getByTitle('Candidates')
        .items.select('Id')
        .filter(`JobOpeningId eq ${jobId} and Email eq '${safe}'`)
        .top(1)();
      return items.length > 0;
    } catch {
      return false;
    }
  }

  public async updateHRFeedback(candidateId: number, feedback: string): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle('Candidates').items.getById(candidateId).update({
        HRFeedback: feedback,
      });
    } catch (err) {
      throw new Error(`Failed to save HR feedback: ${(err as Error).message}`);
    }
  }

  public async createCandidate(candidate: Omit<ICandidate, 'id'>): Promise<number> {
    try {
      const result = await this._sp.web.lists.getByTitle('Candidates').items.add({
        JobOpeningId: candidate.jobOpeningId,
        CandidateName: candidate.candidateName,
        Email: candidate.email,
        Phone: candidate.phone,
        ResumeUrl: candidate.resumeUrl,
        FitmentScore: candidate.fitmentScore,
        MatchingSkills: candidate.matchingSkills,
        MissingSkills: candidate.missingSkills,
        AISummary: candidate.aiSummary,
        HRFeedback: candidate.hrFeedback,
        Recommendation: candidate.recommendation,
        ExperienceMatch: candidate.experienceMatch,
      });
      const candidateId = (result.data as { Id: number }).Id;

      // Referral fields are stored separately; silently skip if columns don't exist yet
      if (candidate.referredBy || candidate.referrerEmployeeId) {
        try {
          await this._sp.web.lists.getByTitle('Candidates').items.getById(candidateId).update({
            ReferredBy: candidate.referredBy ?? '',
            ReferrerEmail: candidate.referrerEmail ?? '',
            ReferrerEmployeeId: candidate.referrerEmployeeId ?? '',
            ReferrerDesignation: candidate.referrerDesignation ?? '',
          });
        } catch {
          // Columns not yet created — referral data not saved to SP, but record exists
        }
      }
      return candidateId;
    } catch (err) {
      throw new Error(`Failed to create candidate record: ${(err as Error).message}`);
    }
  }

  public async updateCandidateFitment(candidateId: number, report: IFitmentReport): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle('Candidates').items.getById(candidateId).update({
        FitmentScore:    report.fitmentScore,
        MatchingSkills:  report.matchingSkills.join(', '),
        MissingSkills:   report.missingSkills.join(', '),
        AISummary:       report.summary,
        Recommendation:  report.recommendation,
        ExperienceMatch: report.experienceMatch,
      });
    } catch (err) {
      throw new Error(`Failed to save fitment report: ${(err as Error).message}`);
    }
  }

  public async rejectCandidate(candidateId: number, restore = false): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle('Candidates').items.getById(candidateId).update({
        ApplicationStatus: restore ? 'Received' : 'Rejected',
      });
    } catch (err) {
      throw new Error(`Failed to update candidate status: ${(err as Error).message}`);
    }
  }

  // Fetches a file from SharePoint and extracts readable text.
  // Uses pdfjs-dist for PDF parsing (handles compressed streams).
  // DOCX is not supported in the browser bundle.
  public async fetchResumeText(rawUrl: string): Promise<string> {
    // Normalize: strip origin if an absolute URL was stored (e.g. from Power Automate)
    let serverRelativeUrl = rawUrl;
    try {
      const parsed = new URL(rawUrl);
      serverRelativeUrl = parsed.pathname + parsed.search;
    } catch {
      // rawUrl is already server-relative — use as-is
    }

    const buffer = await this._sp.web
      .getFileByServerRelativePath(serverRelativeUrl)
      .getBuffer();

    const ext = serverRelativeUrl.split('.').pop()?.toLowerCase() ?? '';

    if (ext === 'txt' || ext === 'html' || ext === 'htm') {
      return new TextDecoder('utf-8').decode(buffer);
    }

    if (ext === 'pdf') {
      // Send the raw PDF to Claude as a document — no browser-side parsing needed.
      // Encoded as a data URI so AIService can detect it and use the document content type.
      const arr = new Uint8Array(buffer);
      let binary = '';
      for (let i = 0; i < arr.length; i++) binary += String.fromCharCode(arr[i]);
      return 'data:application/pdf;base64,' + btoa(binary);
    }

    if (ext === 'docx' || ext === 'doc') {
      throw new Error('DOCX files cannot be parsed in the browser. Please re-upload the resume as .pdf or .txt.');
    }

    const fallback = new TextDecoder('utf-8').decode(buffer);
    if (fallback.trim().length < 50) {
      throw new Error(`Unsupported resume format (.${ext}). Please upload a .pdf or .txt file.`);
    }
    return fallback;
  }

  // ── Interviews ─────────────────────────────────────────────────────────────

  public async getInterviewsByCandidate(candidateId: number): Promise<IInterview[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('Interviews')
        .items.select(
          'Id', 'CandidateId', 'JobOpeningId', 'InterviewRound',
          'InterviewerEmail', 'ScheduledDate', 'FeedbackStatus', 'Feedback', 'HRNotes'
        )
        .filter(`CandidateId eq ${candidateId}`)
        .orderBy('InterviewRound')();

      return items.map(this._mapInterview);
    } catch (err) {
      throw new Error(`Failed to load interviews: ${(err as Error).message}`);
    }
  }

  public async getInterviewsByJobOpening(jobId: number): Promise<IInterview[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('Interviews')
        .items.select(
          'Id', 'CandidateId', 'JobOpeningId', 'InterviewRound',
          'InterviewerEmail', 'ScheduledDate', 'FeedbackStatus', 'Feedback', 'HRNotes'
        )
        .filter(`JobOpeningId eq ${jobId}`)
        .orderBy('InterviewRound')();
      return items.map(this._mapInterview);
    } catch (err) {
      throw new Error(`Failed to load interviews: ${(err as Error).message}`);
    }
  }

  public async getAllInterviews(): Promise<IInterview[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('Interviews')
        .items.select(
          'Id', 'CandidateId', 'JobOpeningId', 'InterviewRound',
          'InterviewerEmail', 'ScheduledDate', 'FeedbackStatus', 'Feedback', 'HRNotes'
        )
        .top(5000)();

      return items.map(this._mapInterview);
    } catch (err) {
      throw new Error(`Failed to load interviews: ${(err as Error).message}`);
    }
  }

  public async getAllPendingInterviews(): Promise<IInterview[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('Interviews')
        .items.select(
          'Id', 'CandidateId', 'JobOpeningId', 'InterviewRound',
          'InterviewerEmail', 'ScheduledDate', 'FeedbackStatus', 'Feedback', 'HRNotes'
        )
        .filter("FeedbackStatus eq 'Pending'")();

      return items.map(this._mapInterview);
    } catch (err) {
      throw new Error(`Failed to load pending interviews: ${(err as Error).message}`);
    }
  }

  public async createInterview(interview: Omit<IInterview, 'id'>): Promise<number> {
    try {
      const result = await this._sp.web.lists.getByTitle('Interviews').items.add({
        CandidateId: interview.candidateId,
        JobOpeningId: interview.jobOpeningId,
        InterviewRound: interview.interviewRound,
        InterviewerEmail: interview.interviewerEmail,
        ScheduledDate: interview.scheduledDate,
        FeedbackStatus: 'Pending',
        Feedback: '',
        HRNotes: interview.hrNotes ?? '',
      });
      return (result.data as { Id: number }).Id;
    } catch (err) {
      throw new Error(`Failed to create interview: ${(err as Error).message}`);
    }
  }

  public async updateHRNotes(interviewId: number, hrNotes: string): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle('Interviews').items.getById(interviewId).update({
        HRNotes: hrNotes,
      });
    } catch (err) {
      throw new Error(`Failed to update HR notes: ${(err as Error).message}`);
    }
  }

  // ── Resumes library folders ────────────────────────────────────────────────

  public async ensureResumeFolder(department: string, jobTitle: string): Promise<string> {
    const sanitize = (s: string): string => s.replace(/[^a-zA-Z0-9_\- ]/g, '').trim();
    const deptFolder = sanitize(department);
    const jtFolder = sanitize(jobTitle);
    const libUrl = await this._getResumesLibraryUrl(); // e.g. "/Resumes"

    // addUsingPath does not create parent folders — create each level explicitly
    const levels = [
      `${libUrl}/${deptFolder}`,
      `${libUrl}/${deptFolder}/${jtFolder}`,
      `${libUrl}/${deptFolder}/${jtFolder}/Direct`,
      `${libUrl}/${deptFolder}/${jtFolder}/Referred`,
    ];

    for (const absPath of levels) {
      try {
        await this._sp.web.folders.addUsingPath(absPath, true);
      } catch {
        // Folder already exists — continue
      }
    }
    return `${deptFolder}/${jtFolder}`;
  }

  // ── JD Template ───────────────────────────────────────────────────────────

  public async getJDTemplateText(): Promise<string> {
    try {
      // Find a file whose name starts with "JD" in the root of the Resumes library
      const items = await this._sp.web.lists
        .getByTitle('Resumes')
        .items.select('FileLeafRef', 'FileRef')
        .filter("startswith(FileLeafRef,'JD')")
        .top(10)();

      type FileItem = { FileLeafRef: string; FileRef: string };

      const match = (items as FileItem[]).find(i => {
        const name = i.FileLeafRef.toLowerCase();
        return name.startsWith('jd template') || name.startsWith('jd_template');
      });

      if (!match) return '';

      const ext = match.FileLeafRef.toLowerCase().split('.').pop() ?? '';

      // DOCX cannot be parsed in the browser bundle (no Node.js fs/path available).
      // Save your JD template as a .txt or .html file to enable template-based generation.
      // DOCX parsing is supported in the Azure Function (server-side) for resume screening.
      if (ext === 'docx' || ext === 'doc') {
        return '';
      }

      const buffer = await this._sp.web.getFileByServerRelativePath(match.FileRef).getBuffer();
      return new TextDecoder('utf-8').decode(buffer);
    } catch {
      return ''; // Template unreadable — Claude generates without it
    }
  }

  // ── Resume upload ──────────────────────────────────────────────────────────

  public async uploadResume(
    department: string,
    jobTitle: string,
    file: File,
    candidateName: string,
    subfolder: 'Referred' | 'Direct' = 'Direct'
  ): Promise<string> {
    await this.ensureResumeFolder(department, jobTitle);
    const sanitize = (s: string): string => s.replace(/[^a-zA-Z0-9_\- ]/g, '').trim();
    const libUrl = await this._getResumesLibraryUrl(); // e.g. "/Resumes"
    const sanitizedName = candidateName
      .replace(/[^a-zA-Z0-9_\- ]/g, '')
      .trim()
      .replace(/ /g, '_');
    const ext = file.name.split('.').pop() ?? 'pdf';
    const fileName = `${sanitizedName}_${Date.now()}.${ext}`;
    const serverRelFolderPath = `${libUrl}/${sanitize(department)}/${sanitize(jobTitle)}/${subfolder}`;

    const buffer = await file.arrayBuffer();
    const result = await this._sp.web
      .getFolderByServerRelativePath(serverRelFolderPath)
      .files.addUsingPath(fileName, buffer, { Overwrite: true });

    return (result.data as { ServerRelativeUrl: string }).ServerRelativeUrl;
  }

  // ── Private helpers ────────────────────────────────────────────────────────

  private async _getWebUrl(): Promise<string> {
    if (!this._webUrl) {
      const info = await this._sp.web.select('ServerRelativeUrl')();
      this._webUrl = (info as unknown as { ServerRelativeUrl: string })
        .ServerRelativeUrl.replace(/\/$/, '');
    }
    return this._webUrl;
  }

  // Returns the server-relative root URL of the Resumes document library,
  // e.g. "/Resumes" for root-site or "/sites/HR/Resumes" for a subsite.
  private async _getResumesLibraryUrl(): Promise<string> {
    if (!this._resumesLibraryUrl) {
      const folder = await this._sp.web.lists
        .getByTitle('Resumes')
        .rootFolder
        .select('ServerRelativeUrl')();
      this._resumesLibraryUrl = (folder as unknown as { ServerRelativeUrl: string })
        .ServerRelativeUrl.replace(/\/$/, '');
    }
    return this._resumesLibraryUrl;
  }

  // ── Private mappers ────────────────────────────────────────────────────────

  private _mapJobOpening(i: Record<string, unknown>): IJobOpening {
    return {
      id: i.Id as number,
      title: (i.Title as string) ?? '',
      department: (i.Department as string) ?? '',
      jobTitle: (i.JobTitle as string) ?? '',
      requiredSkills: (i.RequiredSkills as string) ?? '',
      goodToHaveSkills: (i.GoodToHaveSkills as string) ?? '',
      jobLocation: (i.JobLocation as string) ?? '',
      jobType: (i.JobType as IJobOpening['jobType']) ?? '',
      experience: (i.Experience as string) ?? '',
      dueDate: (i.DueDate as string) ?? '',
      status: (i.Status as IJobOpening['status']) ?? 'Open',
      postedBy: (i.PostedBy as string) ?? '',
      linkedInUrl: (i.LinkedInUrl as string) ?? '',
      jobDescription: (i.JobDescription as string) ?? '',
      applicationFormUrl: (i.ApplicationFormUrl as string) ?? '',
    };
  }

  private _mapCandidate(i: Record<string, unknown>): ICandidate {
    return {
      id: i.Id as number,
      jobOpeningId: (i.JobOpeningId as number) ?? 0,
      candidateName: (i.CandidateName as string) ?? '',
      email: (i.Email as string) ?? '',
      phone: (i.Phone as string) ?? '',
      resumeUrl: (i.ResumeUrl as string) ?? '',
      fitmentScore: (i.FitmentScore as number) ?? 0,
      matchingSkills: (i.MatchingSkills as string) ?? '',
      missingSkills: (i.MissingSkills as string) ?? '',
      aiSummary: (i.AISummary as string) ?? '',
      hrFeedback: (i.HRFeedback as string) ?? '',
      recommendation: (i.Recommendation as ICandidate['recommendation']) ?? '',
      experienceMatch: (i.ExperienceMatch as ICandidate['experienceMatch']) ?? '',
      applicationStatus: (i.ApplicationStatus as string) ?? undefined,
      referredBy: (i.ReferredBy as string) ?? undefined,
      referrerEmail: (i.ReferrerEmail as string) ?? undefined,
      referrerEmployeeId: (i.ReferrerEmployeeId as string) ?? undefined,
      referrerDesignation: (i.ReferrerDesignation as string) ?? undefined,
    };
  }

  private _mapInterview(i: Record<string, unknown>): IInterview {
    return {
      id: i.Id as number,
      candidateId: (i.CandidateId as number) ?? 0,
      jobOpeningId: (i.JobOpeningId as number) ?? 0,
      interviewRound: (i.InterviewRound as IInterview['interviewRound']) ?? '1',
      interviewerEmail: (i.InterviewerEmail as string) ?? '',
      scheduledDate: (i.ScheduledDate as string) ?? '',
      feedbackStatus: (i.FeedbackStatus as IInterview['feedbackStatus']) ?? 'Pending',
      feedback: (i.Feedback as string) ?? '',
      hrNotes: (i.HRNotes as string) ?? '',
    };
  }
}
