import { ICandidate, IInterview } from './models';

export type CandidateCategory = 'Received' | 'Screened' | 'Round 1' | 'Round 2' | 'HR Discussion' | 'Final Discussion' | 'Rejected';

export const CATEGORY_ORDER: CandidateCategory[] = ['Received', 'Screened', 'Round 1', 'Round 2', 'HR Discussion', 'Final Discussion', 'Rejected'];

export const CATEGORY_CONFIG: Record<CandidateCategory, { label: string; color: string }> = {
  'Received':          { label: 'Pending Screening',       color: '#605e5c' },
  'Screened':          { label: 'Screened',                color: '#0078d4' },
  'Round 1':           { label: 'Shortlisted — Round 1',   color: '#8764b8' },
  'Round 2':           { label: 'Shortlisted — Round 2',   color: '#498205' },
  'HR Discussion':     { label: 'HR Discussion',           color: '#ca5010' },
  'Final Discussion':  { label: 'Final Discussion',        color: '#107c10' },
  'Rejected':          { label: 'Rejected',                color: '#a80000' },
};

export function deriveCategory(candidate: ICandidate, jobInterviews: IInterview[]): CandidateCategory {
  if (candidate.applicationStatus === 'Rejected') return 'Rejected';
  const mine = jobInterviews.filter(iv => iv.candidateId === candidate.id);
  const maxRound = mine.reduce((max, iv) => Math.max(max, parseInt(iv.interviewRound, 10)), 0);
  if (maxRound >= 3) {
    const round3Done = mine.some(iv => iv.interviewRound === '3' && iv.feedbackStatus === 'Submitted');
    return round3Done && candidate.recommendation === 'Recommended' ? 'Final Discussion' : 'HR Discussion';
  }
  if (maxRound === 2) return 'Round 2';
  if (maxRound === 1) return 'Round 1';
  return candidate.fitmentScore > 0 ? 'Screened' : 'Received';
}
