import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  ActionButton,
} from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { SpService } from '../shared/SpService';
import { GraphService } from '../shared/GraphService';
import { EmailService } from '../shared/EmailService';
import { IJobOpening, ICandidate } from '../shared/models';
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
  loadingCandidatesFor: Set<number>;
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

  constructor(props: ITrackProgressProps) {
    super(props);
    this._spService = new SpService(props.sp);
    this._emailService = new EmailService(props.graphService);
    this.state = {
      jobs: [],
      loadingJobs: true,
      jobsError: '',
      expandedJobIds: new Set(),
      candidatesMap: {},
      loadingCandidatesFor: new Set(),
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
      const candidates = await this._spService.getCandidatesByJobOpening(jobId);
      this.setState(prev => ({
        candidatesMap: { ...prev.candidatesMap, [jobId]: candidates },
        loadingCandidatesFor: new Set(Array.from(prev.loadingCandidatesFor).filter(id => id !== jobId)),
      }));
    } catch {
      this.setState(prev => ({
        loadingCandidatesFor: new Set(Array.from(prev.loadingCandidatesFor).filter(id => id !== jobId)),
      }));
    }
  }

  private _renderJobCard(job: IJobOpening): React.ReactNode {
    const { expandedJobIds, candidatesMap, loadingCandidatesFor } = this.state;
    const isExpanded = expandedJobIds.has(job.id);
    const candidates = candidatesMap[job.id] ?? [];
    const isLoadingCandidates = loadingCandidatesFor.has(job.id);
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
            ) : candidates.length === 0 ? (
              <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', padding: '12px 0' } }}>
                No candidates found. Upload resumes to the Resumes library to trigger AI screening.
              </Text>
            ) : (
              <>
                <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', padding: '12px 0 0' } }}>
                  {candidates.length} candidate{candidates.length !== 1 ? 's' : ''} — ranked by fitment score
                </Text>
                {candidates.map(c => (
                  <CandidateCard
                    key={c.id}
                    candidate={c}
                    job={job}
                    spService={this._spService}
                    emailService={this._emailService}
                  />
                ))}
              </>
            )}
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

        {!loadingJobs && !jobsError && jobs.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info}>
            No open or in-progress job openings found. Use the Post Job tab to create one.
          </MessageBar>
        )}

        {jobs.map(job => this._renderJobCard(job))}
      </div>
    );
  }
}
