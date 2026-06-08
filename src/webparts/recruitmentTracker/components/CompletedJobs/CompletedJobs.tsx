import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  ActionButton,
  ProgressIndicator,
} from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { SpService } from '../shared/SpService';
import { IJobOpening, ICandidate, IInterview } from '../shared/models';
import styles from './CompletedJobs.module.scss';

interface ICompletedJobsProps {
  sp: SPFI;
}

interface ICompletedJobsState {
  jobs: IJobOpening[];
  loadingJobs: boolean;
  jobsError: string;
  expandedJobIds: Set<number>;
  candidatesMap: Record<number, ICandidate[]>;
  interviewsMap: Record<number, IInterview[]>;
  loadingDataFor: Set<number>;
}

function getScoreColor(score: number): string {
  if (score >= 75) return '#107c10';
  if (score >= 50) return '#f7630c';
  return '#a80000';
}

function getRecommendationClass(recommendation: string): string {
  switch (recommendation) {
    case 'Recommended':     return styles.badgeGreen;
    case 'Maybe':           return styles.badgeOrange;
    case 'Not Recommended': return styles.badgeRed;
    default:                return styles.badgeGrey;
  }
}

function getOutcome(candidate: ICandidate): { label: string; cssClass: string } {
  if (candidate.applicationStatus === 'Rejected') {
    return { label: 'Rejected', cssClass: styles.badgeRed };
  }
  if (candidate.recommendation === 'Recommended') {
    return { label: 'Selected', cssClass: styles.badgeGreen };
  }
  return { label: 'Not Selected', cssClass: styles.badgeGrey };
}

function renderSkillChips(skills: string, chipClass: string): React.ReactNode {
  const list = skills.split(',').map(s => s.trim()).filter(Boolean);
  if (list.length === 0) return null;
  return (
    <div className={styles.chipRow}>
      {list.map(s => (
        <span key={s} className={`${styles.chip} ${chipClass}`}>{s}</span>
      ))}
    </div>
  );
}

function csvEscape(value: string): string {
  if (/[",\n\r]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

function summarizeInterviews(candidateId: number, interviews: IInterview[]): string {
  const mine = interviews
    .filter(iv => iv.candidateId === candidateId)
    .sort((a, b) => parseInt(a.interviewRound, 10) - parseInt(b.interviewRound, 10));

  if (mine.length === 0) return 'No interviews';

  return mine
    .map(iv => {
      const date = iv.scheduledDate ? new Date(iv.scheduledDate).toLocaleDateString('en-GB') : '—';
      return `Round ${iv.interviewRound}: ${iv.interviewerEmail} on ${date} — ${iv.feedbackStatus}`;
    })
    .join('; ');
}

function buildCsv(job: IJobOpening, candidates: ICandidate[], interviews: IInterview[]): string {
  const header = [
    'Candidate Name', 'Email', 'Phone', 'Outcome', 'Application Status',
    'Fitment Score', 'Recommendation', 'Experience Match',
    'Matching Skills', 'Missing Skills', 'AI Summary', 'HR Feedback', 'Interviews',
  ];

  const rows = candidates.map(c => [
    c.candidateName,
    c.email,
    c.phone,
    getOutcome(c).label,
    c.applicationStatus ?? '',
    String(c.fitmentScore),
    c.recommendation,
    c.experienceMatch,
    c.matchingSkills,
    c.missingSkills,
    c.aiSummary,
    c.hrFeedback,
    summarizeInterviews(c.id, interviews),
  ]);

  return [header, ...rows]
    .map(row => row.map(csvEscape).join(','))
    .join('\r\n');
}

function downloadCsv(filename: string, csv: string): void {
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function sanitizeFilenamePart(value: string): string {
  return value.replace(/[\\/:*?"<>|]/g, '-').trim();
}

export class CompletedJobs extends React.Component<ICompletedJobsProps, ICompletedJobsState> {
  private _spService: SpService;

  constructor(props: ICompletedJobsProps) {
    super(props);
    this._spService = new SpService(props.sp);
    this.state = {
      jobs: [],
      loadingJobs: true,
      jobsError: '',
      expandedJobIds: new Set(),
      candidatesMap: {},
      interviewsMap: {},
      loadingDataFor: new Set(),
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadJobs();
  }

  private async _loadJobs(): Promise<void> {
    this.setState({ loadingJobs: true, jobsError: '' });
    try {
      const jobs = await this._spService.getClosedJobOpenings();
      this.setState({ jobs, loadingJobs: false });
    } catch (err) {
      this.setState({ loadingJobs: false, jobsError: (err as Error).message });
    }
  }

  private async _toggleJob(jobId: number): Promise<void> {
    const expanded = new Set(this.state.expandedJobIds);

    if (expanded.has(jobId)) {
      expanded.delete(jobId);
      this.setState({ expandedJobIds: expanded });
      return;
    }

    expanded.add(jobId);
    this.setState({ expandedJobIds: expanded });

    // Only fetch candidates/interviews once per job
    if (this.state.candidatesMap[jobId] !== undefined) return;

    const loading = new Set(this.state.loadingDataFor);
    loading.add(jobId);
    this.setState({ loadingDataFor: loading });

    try {
      const [candidates, interviews] = await Promise.all([
        this._spService.getCandidatesByJobOpening(jobId),
        this._spService.getInterviewsByJobOpening(jobId),
      ]);
      this.setState(prev => ({
        candidatesMap: { ...prev.candidatesMap, [jobId]: candidates },
        interviewsMap: { ...prev.interviewsMap, [jobId]: interviews },
        loadingDataFor: new Set(Array.from(prev.loadingDataFor).filter(id => id !== jobId)),
      }));
    } catch {
      this.setState(prev => ({
        loadingDataFor: new Set(Array.from(prev.loadingDataFor).filter(id => id !== jobId)),
      }));
    }
  }

  private _onExportCsv = async (job: IJobOpening): Promise<void> => {
    const cachedCandidates = this.state.candidatesMap[job.id];
    const cachedInterviews = this.state.interviewsMap[job.id];

    let candidates: ICandidate[];
    let interviews: IInterview[];

    if (cachedCandidates === undefined || cachedInterviews === undefined) {
      const fetched = await Promise.all([
        this._spService.getCandidatesByJobOpening(job.id),
        this._spService.getInterviewsByJobOpening(job.id),
      ]);
      candidates = fetched[0];
      interviews = fetched[1];
      this.setState(prev => ({
        candidatesMap: { ...prev.candidatesMap, [job.id]: candidates },
        interviewsMap: { ...prev.interviewsMap, [job.id]: interviews },
      }));
    } else {
      candidates = cachedCandidates;
      interviews = cachedInterviews;
    }

    const csv = buildCsv(job, candidates, interviews);
    const filename = `${sanitizeFilenamePart(job.title)}-candidates.csv`;
    downloadCsv(filename, csv);
  };

  private _renderCandidateRow(candidate: ICandidate, interviews: IInterview[]): React.ReactNode {
    const outcome = getOutcome(candidate);
    const mine = interviews
      .filter(iv => iv.candidateId === candidate.id)
      .sort((a, b) => parseInt(a.interviewRound, 10) - parseInt(b.interviewRound, 10));

    return (
      <div key={candidate.id} className={styles.candidateRow}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 8 }}>
          <Stack tokens={{ childrenGap: 2 }} styles={{ root: { flex: 1, minWidth: 180 } }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              {candidate.candidateName}
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              {candidate.email}{candidate.phone ? ` · ${candidate.phone}` : ''}
            </Text>
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 6 }}>
            <span className={`${styles.badge} ${outcome.cssClass}`}>{outcome.label}</span>
            <span className={`${styles.badge} ${getRecommendationClass(candidate.recommendation)}`}>
              {candidate.recommendation || 'Pending'}
            </span>
          </Stack>
        </Stack>

        <div className={styles.scoreRow}>
          <Text
            variant="small"
            styles={{ root: { color: getScoreColor(candidate.fitmentScore), fontWeight: 600, minWidth: 44 } }}
          >
            {candidate.fitmentScore}%
          </Text>
          <div style={{ flex: 1 }}>
            <ProgressIndicator
              percentComplete={candidate.fitmentScore / 100}
              barHeight={8}
              styles={{
                itemProgress: { padding: 0 },
                progressBar: { backgroundColor: getScoreColor(candidate.fitmentScore) },
              }}
            />
          </div>
          {candidate.experienceMatch && (
            <span className={`${styles.chip} ${
              candidate.experienceMatch === 'meets'   ? styles.chipGreen :
              candidate.experienceMatch === 'exceeds' ? styles.chipBlue  :
              styles.chipRed
            }`}>
              Exp: {candidate.experienceMatch}
            </span>
          )}
        </div>

        {candidate.matchingSkills && renderSkillChips(candidate.matchingSkills, styles.chipGreen)}
        {candidate.missingSkills && renderSkillChips(candidate.missingSkills, styles.chipRed)}

        {candidate.aiSummary && (
          <div className={styles.aiSummary}>
            <Text variant="xSmall" styles={{ root: { color: '#0078d4', fontWeight: 600 } }}>AI Summary</Text>
            <Text variant="small">{candidate.aiSummary}</Text>
          </div>
        )}

        {candidate.hrFeedback && (
          <div style={{ marginTop: 8 }}>
            <Text variant="xSmall" styles={{ root: { color: '#605e5c', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' } }}>
              HR Feedback
            </Text>
            <Text variant="small">{candidate.hrFeedback}</Text>
          </div>
        )}

        {mine.length > 0 && (
          <div className={styles.interviewHistory}>
            <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
              Interview History
            </Text>
            {mine.map(iv => (
              <div key={iv.id} className={styles.interviewRow}>
                <span className={styles.roundBadge}>Rd {iv.interviewRound}</span>
                <Text variant="small" styles={{ root: { flex: 1 } }}>{iv.interviewerEmail}</Text>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  {iv.scheduledDate ? new Date(iv.scheduledDate).toLocaleString('en-GB') : '—'}
                </Text>
                <span className={`${styles.chip} ${iv.feedbackStatus === 'Submitted' ? styles.chipGreen : styles.chipGrey}`}>
                  {iv.feedbackStatus}
                </span>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  private _renderJobBody(job: IJobOpening): React.ReactNode {
    const { candidatesMap, interviewsMap, loadingDataFor } = this.state;
    const candidates = candidatesMap[job.id] ?? [];
    const interviews = interviewsMap[job.id] ?? [];
    const isLoading = loadingDataFor.has(job.id);

    if (isLoading) {
      return <Spinner size={SpinnerSize.medium} label="Loading candidates…" styles={{ root: { padding: 16 } }} />;
    }

    return (
      <>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { padding: '12px 0 4px' } }}>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            {candidates.length} candidate{candidates.length !== 1 ? 's' : ''}
          </Text>
          <ActionButton
            iconProps={{ iconName: 'ExcelDocument' }}
            text="Export CSV"
            onClick={() => { this._onExportCsv(job).catch(() => undefined); }}
            disabled={candidates.length === 0}
          />
        </Stack>

        {candidates.length === 0 ? (
          <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', padding: '4px 0 12px' } }}>
            No candidates were recorded for this job opening.
          </Text>
        ) : (
          candidates.map(c => this._renderCandidateRow(c, interviews))
        )}
      </>
    );
  }

  private _renderJobCard(job: IJobOpening): React.ReactNode {
    const isExpanded = this.state.expandedJobIds.has(job.id);

    return (
      <div key={job.id} className={styles.jobCard}>
        <div
          className={styles.jobCardHeader}
          role="button"
          tabIndex={0}
          onClick={() => { this._toggleJob(job.id).catch(() => undefined); }}
          onKeyDown={e => {
            if (e.key === 'Enter' || e.key === ' ') {
              e.preventDefault();
              this._toggleJob(job.id).catch(() => undefined);
            }
          }}
        >
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} wrap>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, flex: 1, minWidth: 160 } }}>
              {job.title}
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{job.department}</Text>
            <span className={styles.statusBadge}>{job.status}</span>
            {job.dueDate && (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Closed by: {new Date(job.dueDate).toLocaleDateString('en-GB')}
              </Text>
            )}
            <ActionButton
              iconProps={{ iconName: isExpanded ? 'ChevronUp' : 'ChevronDown' }}
              styles={{ root: { padding: 0, height: 'auto', minWidth: 'auto' } }}
              ariaLabel={isExpanded ? 'Collapse' : 'Expand'}
            />
          </Stack>
        </div>

        {isExpanded && (
          <div className={styles.jobCardBody}>
            {this._renderJobBody(job)}
          </div>
        )}
      </div>
    );
  }

  public render(): React.ReactElement {
    const { jobs, loadingJobs, jobsError } = this.state;

    return (
      <div className={styles.container}>
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{ root: { marginBottom: 16 } }}
        >
          <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>Completed Jobs</Text>
          <ActionButton
            iconProps={{ iconName: 'Refresh' }}
            text="Refresh"
            onClick={() => { this._loadJobs().catch(() => undefined); }}
            disabled={loadingJobs}
          />
        </Stack>

        {loadingJobs && <Spinner size={SpinnerSize.medium} label="Loading closed job openings…" />}

        {jobsError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ jobsError: '' })}
          >
            {jobsError}
          </MessageBar>
        )}

        {!loadingJobs && !jobsError && jobs.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info}>
            No closed job openings yet. Jobs appear here once their status is set to &quot;Closed&quot;.
          </MessageBar>
        )}

        {jobs.map(job => this._renderJobCard(job))}
      </div>
    );
  }
}
