import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  ActionButton,
  DefaultButton,
  PrimaryButton,
  Dialog,
  DialogType,
  DialogFooter,
  Panel,
  PanelType,
} from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { SpService } from '../shared/SpService';
import { GraphService } from '../shared/GraphService';
import { EmailService } from '../shared/EmailService';
import { AIService } from '../shared/AIService';
import { IJobOpening, ICandidate, IInterview } from '../shared/models';
import { CATEGORY_ORDER, CATEGORY_CONFIG, deriveCategory } from '../shared/candidateCategory';
import { CandidateCard } from './CandidateCard';
import styles from './TrackProgress.module.scss';

interface ITrackProgressProps {
  sp: SPFI;
  graphService: GraphService;
}

interface ITrackProgressState {
  jobs: IJobOpening[];
  loadingJobs: boolean;
  jobsError: string;
  expandedJobIds: Set<number>;
  candidatesMap: Record<number, ICandidate[]>;
  interviewsMap: Record<number, IInterview[]>;
  loadingCandidatesFor: Set<number>;
  confirmCloseJobId: number | undefined;
  closingJobId: number | undefined;
  closeError: string;
  reportJobId: number | undefined;
  loadingReport: boolean;
}

function daysRemaining(dueDateIso: string): number {
  const due = new Date(dueDateIso);
  const now = new Date();
  return Math.ceil((due.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
}

function getStatusClass(status: IJobOpening['status']): string {
  switch (status) {
    case 'Open':        return styles.statusOpen;
    case 'In Progress': return styles.statusInProgress;
    case 'Closed':      return styles.statusClosed;
    default:            return styles.statusOpen;
  }
}

export class TrackProgress extends React.Component<ITrackProgressProps, ITrackProgressState> {
  private _spService: SpService;
  private _emailService: EmailService;
  private _aiService: AIService;

  constructor(props: ITrackProgressProps) {
    super(props);
    this._spService = new SpService(props.sp);
    this._emailService = new EmailService(props.graphService);
    this._aiService = new AIService();
    this.state = {
      jobs: [],
      loadingJobs: true,
      jobsError: '',
      expandedJobIds: new Set(),
      candidatesMap: {},
      interviewsMap: {},
      loadingCandidatesFor: new Set(),
      confirmCloseJobId: undefined,
      closingJobId: undefined,
      closeError: '',
      reportJobId: undefined,
      loadingReport: false,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadJobs();
  }

  private async _loadJobs(): Promise<void> {
    this.setState({ loadingJobs: true, jobsError: '' });
    try {
      const jobs = await this._spService.getOpenJobOpenings();
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

    // Only fetch candidates once per job
    if (this.state.candidatesMap[jobId] !== undefined) return;

    const loading = new Set(this.state.loadingCandidatesFor);
    loading.add(jobId);
    this.setState({ loadingCandidatesFor: loading });

    try {
      const [candidates, interviews] = await Promise.all([
        this._spService.getCandidatesByJobOpening(jobId),
        this._spService.getInterviewsByJobOpening(jobId),
      ]);
      this.setState(prev => ({
        candidatesMap: { ...prev.candidatesMap, [jobId]: candidates },
        interviewsMap: { ...prev.interviewsMap, [jobId]: interviews },
        loadingCandidatesFor: new Set(Array.from(prev.loadingCandidatesFor).filter(id => id !== jobId)),
      }));
    } catch {
      this.setState(prev => ({
        loadingCandidatesFor: new Set(Array.from(prev.loadingCandidatesFor).filter(id => id !== jobId)),
      }));
    }
  }

  private _onCandidateUpdated = (updated: ICandidate): void => {
    this.setState(prev => {
      const list = prev.candidatesMap[updated.jobOpeningId] ?? [];
      return {
        candidatesMap: {
          ...prev.candidatesMap,
          [updated.jobOpeningId]: list.map(c => c.id === updated.id ? updated : c),
        },
      };
    });
  };

  private _refreshInterviewsForJob = (jobId: number): void => {
    this._spService.getInterviewsByJobOpening(jobId)
      .then(interviews => {
        this.setState(prev => ({
          interviewsMap: { ...prev.interviewsMap, [jobId]: interviews },
        }));
      })
      .catch(() => undefined);
  };

  private _onConfirmCloseJob = async (): Promise<void> => {
    const jobId = this.state.confirmCloseJobId;
    if (jobId === undefined) return;

    this.setState({ confirmCloseJobId: undefined, closingJobId: jobId, closeError: '' });
    try {
      await this._spService.updateJobOpeningStatus(jobId, 'Closed');
      this.setState(prev => ({
        jobs: prev.jobs.filter(j => j.id !== jobId),
        expandedJobIds: new Set(Array.from(prev.expandedJobIds).filter(id => id !== jobId)),
        closingJobId: undefined,
      }));
    } catch (err) {
      this.setState({ closingJobId: undefined, closeError: (err as Error).message });
    }
  };

  private _openReport = (jobId: number): void => {
    this.setState({ reportJobId: jobId });
    if (this.state.candidatesMap[jobId] !== undefined) return;

    this.setState({ loadingReport: true });
    Promise.all([
      this._spService.getCandidatesByJobOpening(jobId),
      this._spService.getInterviewsByJobOpening(jobId),
    ])
      .then(fetched => {
        const candidates = fetched[0];
        const interviews = fetched[1];
        this.setState(prev => ({
          candidatesMap: { ...prev.candidatesMap, [jobId]: candidates },
          interviewsMap: { ...prev.interviewsMap, [jobId]: interviews },
          loadingReport: false,
        }));
      })
      .catch(() => this.setState({ loadingReport: false }));
  };

  private _renderReportPanel(): React.ReactNode {
    const { reportJobId, jobs, candidatesMap, interviewsMap, loadingReport } = this.state;
    if (reportJobId === undefined) return null;

    const job = jobs.find(j => j.id === reportJobId);
    if (!job) return null;

    const candidates = candidatesMap[reportJobId] ?? [];
    const interviews = interviewsMap[reportJobId] ?? [];

    const scoreColor = (score: number): string =>
      score >= 75 ? '#107c10' : score >= 50 ? '#f7630c' : score > 0 ? '#a80000' : '#605e5c';
    const recBg = (r: string): string =>
      r === 'Recommended' ? '#dff6dd' : r === 'Maybe' ? '#fff4ce' : r === 'Not Recommended' ? '#fde7e9' : '#edebe9';
    const recFg = (r: string): string =>
      r === 'Recommended' ? '#107c10' : r === 'Maybe' ? '#8a8000' : r === 'Not Recommended' ? '#a80000' : '#605e5c';

    // Summary counts per category
    const categoryCounts: Partial<Record<string, number>> = {};
    candidates.forEach(c => {
      const cat = deriveCategory(c, interviews);
      categoryCounts[cat] = (categoryCounts[cat] ?? 0) + 1;
    });

    return (
      <Panel
        isOpen
        type={PanelType.medium}
        headerText={`Pipeline Report — ${job.title}`}
        onDismiss={() => this.setState({ reportJobId: undefined })}
      >
        <div style={{ padding: '16px 0 24px' }}>
          {/* Job metadata */}
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 16 }}>
            <span style={{ background: '#edebe9', color: '#323130', padding: '3px 10px', borderRadius: 12, fontSize: 12 }}>{job.department}</span>
            {job.jobLocation && (
              <span style={{ background: '#edebe9', color: '#323130', padding: '3px 10px', borderRadius: 12, fontSize: 12 }}>{job.jobLocation}</span>
            )}
            {job.experience && (
              <span style={{ background: '#edebe9', color: '#323130', padding: '3px 10px', borderRadius: 12, fontSize: 12 }}>{job.experience}</span>
            )}
            {job.dueDate && (
              <span style={{ background: '#fff4ce', color: '#8a8000', padding: '3px 10px', borderRadius: 12, fontSize: 12 }}>
                Due {new Date(job.dueDate).toLocaleDateString('en-GB')}
              </span>
            )}
          </div>

          {loadingReport && (
            <Spinner size={SpinnerSize.medium} label="Loading candidates…" />
          )}

          {!loadingReport && candidates.length === 0 && (
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>No candidates found for this job.</Text>
          )}

          {!loadingReport && candidates.length > 0 && (
            <>
              {/* Summary row */}
              <div style={{ background: '#f3f2f1', borderRadius: 6, padding: '12px 14px', marginBottom: 20 }}>
                <Text variant="small" styles={{ root: { fontWeight: 600, display: 'block', marginBottom: 8 } }}>
                  Total: {candidates.length} candidate{candidates.length !== 1 ? 's' : ''}
                </Text>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                  {CATEGORY_ORDER.map(cat => {
                    const count = categoryCounts[cat];
                    if (!count) return null;
                    const { label, color } = CATEGORY_CONFIG[cat];
                    return (
                      <span
                        key={cat}
                        style={{ background: color, color: '#fff', padding: '2px 10px', borderRadius: 10, fontSize: 12, fontWeight: 600 }}
                      >
                        {label}: {count}
                      </span>
                    );
                  })}
                </div>
              </div>

              {/* Pipeline breakdown */}
              {CATEGORY_ORDER.map(cat => {
                const group = candidates.filter(c => deriveCategory(c, interviews) === cat);
                if (group.length === 0) return null;
                const { label, color } = CATEGORY_CONFIG[cat];
                return (
                  <div key={cat} style={{ marginBottom: 20 }}>
                    {/* Stage header */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, borderBottom: `2px solid ${color}`, paddingBottom: 6, marginBottom: 8 }}>
                      <span style={{ fontSize: 13, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.4px', color }}>{label}</span>
                      <span style={{ background: color, color: '#fff', borderRadius: 10, padding: '1px 8px', fontSize: 11, fontWeight: 700 }}>{group.length}</span>
                    </div>

                    {/* Candidate rows */}
                    {group.map(c => {
                      const roundNum = cat === 'Round 1' ? '1' : cat === 'Round 2' ? '2' : undefined;
                      const interview = roundNum
                        ? interviews.find(iv => iv.candidateId === c.id && iv.interviewRound === roundNum)
                        : undefined;

                      return (
                        <div
                          key={c.id}
                          style={{ padding: '8px 10px', borderRadius: 4, marginBottom: 4, background: '#faf9f8', border: '1px solid #edebe9' }}
                        >
                          {/* Name + email + badges */}
                          <div style={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
                            <div style={{ flex: 1, minWidth: 120 }}>
                              <div style={{ fontWeight: 600, fontSize: 13, color: '#323130' }}>{c.candidateName}</div>
                              <div style={{ fontSize: 11, color: '#605e5c' }}>{c.email}</div>
                            </div>
                            <span style={{
                              background: scoreColor(c.fitmentScore), color: '#fff',
                              padding: '2px 10px', borderRadius: 10, fontSize: 12, fontWeight: 700,
                              minWidth: 44, textAlign: 'center',
                            }}>
                              {c.fitmentScore > 0 ? `${c.fitmentScore}%` : '—'}
                            </span>
                            {c.recommendation && (
                              <span style={{ background: recBg(c.recommendation), color: recFg(c.recommendation), padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600 }}>
                                {c.recommendation}
                              </span>
                            )}
                          </div>

                          {/* Interview date — Round 1 / Round 2 */}
                          {interview && interview.scheduledDate && (
                            <div style={{ marginTop: 6, display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 6, fontSize: 12 }}>
                              <span style={{ color: '#605e5c' }}>Interview:</span>
                              <strong style={{ color: '#0078d4' }}>
                                {new Date(interview.scheduledDate).toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })}
                              </strong>
                              <span style={{ color: '#605e5c' }}>· {interview.interviewerEmail}</span>
                              <span style={{
                                background: interview.feedbackStatus === 'Submitted' ? '#dff6dd' : '#fff4ce',
                                color: interview.feedbackStatus === 'Submitted' ? '#107c10' : '#8a8000',
                                padding: '1px 7px', borderRadius: 8, fontSize: 11, fontWeight: 600,
                              }}>
                                {interview.feedbackStatus}
                              </span>
                            </div>
                          )}

                          {/* HR feedback — Rejected candidates only */}
                          {cat === 'Rejected' && c.hrFeedback && (
                            <div style={{
                              marginTop: 6, padding: '5px 8px',
                              background: '#fde7e9', borderLeft: '3px solid #a80000',
                              borderRadius: '0 3px 3px 0', fontSize: 12, color: '#323130',
                            }}>
                              <span style={{ fontSize: 11, fontWeight: 600, color: '#a80000', textTransform: 'uppercase', letterSpacing: '0.4px' }}>
                                HR Feedback:{' '}
                              </span>
                              {c.hrFeedback}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                );
              })}
            </>
          )}
        </div>
      </Panel>
    );
  }

  private _renderCandidateGroups(job: IJobOpening, candidates: ICandidate[], interviews: IInterview[]): React.ReactNode {
    if (candidates.length === 0) {
      return (
        <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', padding: '12px 0' } }}>
          No candidates yet. Use the Ongoing Positions tab to refer or add candidates.
        </Text>
      );
    }

    return (
      <>
        <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', padding: '12px 0 0' } }}>
          {candidates.length} candidate{candidates.length !== 1 ? 's' : ''}
        </Text>
        {CATEGORY_ORDER.map(cat => {
          const group = candidates.filter(c => deriveCategory(c, interviews) === cat);
          if (group.length === 0) return null;
          const { label, color } = CATEGORY_CONFIG[cat];
          return (
            <div key={cat} className={styles.categorySection}>
              <div className={styles.categoryHeader} style={{ borderColor: color, color }}>
                <span className={styles.categoryLabel} style={{ color }}>{label}</span>
                <span className={styles.categoryCount} style={{ background: color }}>
                  <span className={styles.categoryCountText}>{group.length}</span>
                </span>
              </div>
              {group.map(c => (
                <CandidateCard
                  key={c.id}
                  candidate={c}
                  job={job}
                  spService={this._spService}
                  emailService={this._emailService}
                  aiService={this._aiService}
                  onCandidateUpdated={this._onCandidateUpdated}
                  onInterviewScheduled={() => this._refreshInterviewsForJob(job.id)}
                />
              ))}
            </div>
          );
        })}
      </>
    );
  }

  private _renderJobCard(job: IJobOpening): React.ReactNode {
    const { expandedJobIds, candidatesMap, interviewsMap, loadingCandidatesFor, closingJobId } = this.state;
    const isExpanded = expandedJobIds.has(job.id);
    const candidates = candidatesMap[job.id] ?? [];
    const interviews = interviewsMap[job.id] ?? [];
    const isLoadingCandidates = loadingCandidatesFor.has(job.id);
    const isClosing = closingJobId === job.id;
    const days = job.dueDate ? daysRemaining(job.dueDate) : null;

    return (
      <div key={job.id} className={styles.jobCard}>
        {/* Clickable header */}
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
            <span className={`${styles.statusBadge} ${getStatusClass(job.status)}`}>{job.status}</span>
            {job.dueDate && (
              <Text variant="small" styles={{ root: { color: days !== null && days < 7 ? '#a80000' : '#605e5c' } }}>
                Due: {new Date(job.dueDate).toLocaleDateString('en-GB')}
                {days !== null && ` (${days >= 0 ? `${days}d left` : 'Overdue'})`}
              </Text>
            )}
            <DefaultButton
              text="Report"
              iconProps={{ iconName: 'ReportDocument' }}
              onClick={e => {
                e.stopPropagation();
                this._openReport(job.id);
              }}
            />
            <DefaultButton
              text={isClosing ? 'Closing…' : 'Close Job'}
              iconProps={{ iconName: 'Completed' }}
              onClick={e => {
                e.stopPropagation();
                this.setState({ confirmCloseJobId: job.id });
              }}
              disabled={isClosing}
            />
            {isClosing && <Spinner size={SpinnerSize.small} />}
            <ActionButton
              iconProps={{ iconName: isExpanded ? 'ChevronUp' : 'ChevronDown' }}
              styles={{ root: { padding: 0, height: 'auto', minWidth: 'auto' } }}
              ariaLabel={isExpanded ? 'Collapse' : 'Expand'}
            />
          </Stack>
        </div>

        {/* Expanded body */}
        {isExpanded && (
          <div className={styles.jobCardBody}>
            {isLoadingCandidates ? (
              <Spinner size={SpinnerSize.medium} label="Loading candidates…" styles={{ root: { padding: 16 } }} />
            ) : (
              this._renderCandidateGroups(job, candidates, interviews)
            )}
          </div>
        )}
      </div>
    );
  }

  public render(): React.ReactElement {
    const { jobs, loadingJobs, jobsError, confirmCloseJobId, closeError } = this.state;
    const confirmCloseJob = confirmCloseJobId !== undefined ? jobs.find(j => j.id === confirmCloseJobId) : undefined;

    return (
      <div className={styles.container}>
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{ root: { marginBottom: 16 } }}
        >
          <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>Track Progress</Text>
          <ActionButton
            iconProps={{ iconName: 'Refresh' }}
            text="Refresh"
            onClick={() => { this._loadJobs().catch(() => undefined); }}
            disabled={loadingJobs}
          />
        </Stack>

        {loadingJobs && <Spinner size={SpinnerSize.medium} label="Loading job openings…" />}

        {jobsError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ jobsError: '' })}
          >
            {jobsError}
          </MessageBar>
        )}

        {closeError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ closeError: '' })}
          >
            {closeError}
          </MessageBar>
        )}

        {!loadingJobs && !jobsError && jobs.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info}>
            No open or in-progress job openings found. Use the Post Job tab to create one.
          </MessageBar>
        )}

        {jobs.map(job => this._renderJobCard(job))}

        {this._renderReportPanel()}

        {confirmCloseJob && (
          <Dialog
            hidden={false}
            onDismiss={() => this.setState({ confirmCloseJobId: undefined })}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Close this job opening?',
              subText: `"${confirmCloseJob.title}" will be marked as Closed, removed from Track Progress, and moved to the Completed Jobs tab. Candidates and interview records are kept and remain visible there.`,
            }}
          >
            <DialogFooter>
              <PrimaryButton text="Close Job" onClick={() => { this._onConfirmCloseJob().catch(() => undefined); }} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ confirmCloseJobId: undefined })} />
            </DialogFooter>
          </Dialog>
        )}
      </div>
    );
  }
}
