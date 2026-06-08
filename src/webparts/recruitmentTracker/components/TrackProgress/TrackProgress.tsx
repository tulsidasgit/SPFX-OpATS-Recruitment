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
