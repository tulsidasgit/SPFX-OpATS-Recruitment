import { getSpClient, getResumeFolderPath } from '../shared/spClient';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';

export interface IApplicationPayload {
  jobId: number;
  // Required fields
  name: string;
  email: string;
  phone: string;
  currentCtc: string;
  expectedCtc: string;
  noticePeriod: string;
  resumeBase64: string;
  resumeFileName: string;
  // Optional fields
  gender?: string;
  reasonForLeaving?: string;
  workCultureExpectation?: string;
  linkedInUrl?: string;
}

export interface ISubmitResult {
  candidateId: number;
  resumeUrl: string;
}

export async function submitApplication(payload: IApplicationPayload): Promise<ISubmitResult> {
  const sp = getSpClient();

  // 1 ── Fetch job opening to get department + jobTitle for folder path
  const jobs = await sp.web.lists
    .getByTitle('JobOpenings')
    .items.select('Id', 'Department', 'JobTitle', 'Status')
    .filter(`Id eq ${payload.jobId}`)
    .top(1)();

  if (jobs.length === 0) throw new Error(`Job opening #${payload.jobId} not found.`);

  const job = jobs[0] as Record<string, string>;
  if (job['Status'] === 'Closed') throw new Error('This job opening is no longer accepting applications.');

  const department = job['Department'] ?? 'General';
  const jobTitle   = job['JobTitle']   ?? 'Role';

  // 2 ── Ensure resume folder exists and upload file
  const folderPath = getResumeFolderPath(department, jobTitle);
  const sp2 = getSpClient();

  try {
    await sp2.web.folders.addUsingPath(folderPath, true);
  } catch {
    // Folder may already exist
  }

  const resumeBuffer = Buffer.from(payload.resumeBase64, 'base64');
  const safeFileName = payload.resumeFileName.replace(/[^a-zA-Z0-9._\- ]/g, '_');
  const timestampedName = `${Date.now()}_${safeFileName}`;

  const uploadResult = await sp2.web
    .getFolderByServerRelativePath(folderPath)
    .files.addUsingPath(timestampedName, resumeBuffer, { Overwrite: false });

  const resumeUrl = (uploadResult.data as { ServerRelativeUrl?: string }).ServerRelativeUrl ?? '';

  // 3 ── Create Candidates list item
  const result = await sp2.web.lists.getByTitle('Candidates').items.add({
    JobOpeningId:            payload.jobId,
    CandidateName:           payload.name,
    Email:                   payload.email,
    Phone:                   payload.phone,
    ResumeUrl:               resumeUrl,
    CurrentCtc:              payload.currentCtc,
    ExpectedCtc:             payload.expectedCtc,
    NoticePeriod:            payload.noticePeriod,
    Gender:                  payload.gender              ?? '',
    ReasonForLeaving:        payload.reasonForLeaving    ?? '',
    WorkCultureExpectation:  payload.workCultureExpectation ?? '',
    LinkedInProfileUrl:      payload.linkedInUrl         ?? '',
    FitmentScore:            0,
    MatchingSkills:          '',
    MissingSkills:           '',
    AISummary:               '',
    HRFeedback:              '',
    ApplicationStatus:       'Received',
  });

  const candidateId = (result.data as { Id: number }).Id;

  // 4 ── Update JobOpening status to "In Progress" if still Open
  if (job['Status'] === 'Open') {
    await sp2.web.lists.getByTitle('JobOpenings').items.getById(payload.jobId).update({
      Status: 'In Progress',
    });
  }

  return { candidateId, resumeUrl };
}
