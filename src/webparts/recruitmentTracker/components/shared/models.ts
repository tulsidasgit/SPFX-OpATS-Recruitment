// ── Domain models for OpATS — Recruitment & Application Tracking System ──

export interface IDepartment {
  id: number;
  title: string;
  isActive: boolean;
}

export interface IJobTitle {
  id: number;
  title: string;
  departmentId: number;
}

export interface IJobOpening {
  id: number;
  title: string;
  department: string;
  jobTitle: string;
  requiredSkills: string;       // Must Have Skills
  goodToHaveSkills: string;     // Good To Have Skills
  jobLocation: string;
  jobType: 'On Site' | 'Hybrid' | 'Remote' | '';
  experience: string;
  dueDate: string;
  status: 'Open' | 'In Progress' | 'Closed';
  postedBy: string;
  linkedInUrl?: string;
  jobDescription?: string;
  applicationFormUrl?: string;
}

export interface ICandidate {
  id: number;
  jobOpeningId: number;
  candidateName: string;
  email: string;
  phone: string;
  resumeUrl: string;
  fitmentScore: number;
  matchingSkills: string;
  missingSkills: string;
  aiSummary: string;
  hrFeedback: string;
  recommendation: 'Recommended' | 'Maybe' | 'Not Recommended' | '';
  experienceMatch: 'meets' | 'exceeds' | 'below' | '';
  applicationStatus?: string;
  referredBy?: string;
  referrerEmail?: string;
  referrerEmployeeId?: string;
  referrerDesignation?: string;
}

export interface IInterview {
  id: number;
  candidateId: number;
  jobOpeningId: number;
  interviewRound: '1' | '2' | '3';
  interviewerEmail: string;
  scheduledDate: string;
  feedbackStatus: 'Pending' | 'Submitted';
  feedback: string;
  hrNotes: string;
}

export interface IFitmentReport {
  candidateName: string;
  contactEmail: string;
  contactPhone: string;
  matchingSkills: string[];
  missingSkills: string[];
  fitmentScore: number;
  experienceMatch: 'meets' | 'exceeds' | 'below';
  summary: string;
  recommendation: 'Recommended' | 'Maybe' | 'Not Recommended';
}

export interface ICurrentUser {
  displayName: string;
  email: string;
  jobTitle?: string;
  employeeId?: string;
}

export interface IEmailNotification {
  to: string[];
  cc?: string[];
  subject: string;
  bodyHtml: string;
}

