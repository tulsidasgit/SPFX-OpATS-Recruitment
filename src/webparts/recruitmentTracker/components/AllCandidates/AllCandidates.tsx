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
import { ICandidate, IInterview, IJobOpening } from '../shared/models';
import { CandidateCategory, CATEGORY_ORDER, CATEGORY_CONFIG, deriveCategory } from '../shared/candidateCategory';
import styles from './AllCandidates.module.scss';

interface IAllCandidatesProps {
  sp: SPFI;
}

interface IAllCandidatesState {
  candidates: ICandidate[];
  interviews: IInterview[];
  jobsById: Record<number, IJobOpening>;
  loading: boolean;
  error: string;
  expandedCategories: Set<CandidateCategory>;
}

function getScoreColor(score: number): string {
  if (score >= 75) return '#107c10';
  if (score >= 50) return '#f7630c';
  return '#a80000';
}

function getRecommendationClass(recommendation: string): string {
  switch (recommendation) {
    case 'Recommended':     return styles.chipGreen;
    case 'Maybe':           return styles.chipBlue;
    case 'Not Recommended': return styles.chipRed;
    default:                return styles.chipGrey;
  }
}

export class AllCandidates extends React.Component<IAllCandidatesProps, IAllCandidatesState> {
  private _spService: SpService;

  constructor(props: IAllCandidatesProps) {
    super(props);
    this._spService = new SpService(props.sp);
    this.state = {
      candidates: [],
      interviews: [],
      jobsById: {},
      loading: true,
      error: '',
      expandedCategories: new Set(),
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadData();
  }

  private _loadData = async (): Promise<void> => {
    this.setState({ loading: true, error: '' });
    try {
      const [candidates, interviews, openJobs, closedJobs] = await Promise.all([
        this._spService.getAllCandidates(),
        this._spService.getAllInterviews(),
        this._spService.getOpenJobOpenings(),
        this._spService.getClosedJobOpenings(),
      ]);

      const jobsById: Record<number, IJobOpening> = {};
      [...openJobs, ...closedJobs].forEach(job => { jobsById[job.id] = job; });

      this.setState({ candidates, interviews, jobsById, loading: false });
    } catch (err) {
      this.setState({ loading: false, error: (err as Error).message });
    }
  };

  private _toggleCategory = (category: CandidateCategory): void => {
    const expanded = new Set(this.state.expandedCategories);
    if (expanded.has(category)) {
      expanded.delete(category);
    } else {
      expanded.add(category);
    }
    this.setState({ expandedCategories: expanded });
  };

  private _renderCandidateRow(candidate: ICandidate): React.ReactNode {
    const job = this.state.jobsById[candidate.jobOpeningId];

    return (
      <div key={candidate.id} className={styles.candidateRow}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 8 }}>
          <Stack tokens={{ childrenGap: 2 }} styles={{ root: { flex: 1, minWidth: 200 } }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              {candidate.candidateName}
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              {candidate.email}{candidate.phone ? ` · ${candidate.phone}` : ''}
            </Text>
          </Stack>
          <Stack tokens={{ childrenGap: 2 }} horizontalAlign="end">
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
              {job ? job.title : `Job #${candidate.jobOpeningId}`}
            </Text>
            {job && (
              <span className={styles.badge}>{job.department}{job.status === 'Closed' ? ' · Closed' : ''}</span>
            )}
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
          {candidate.recommendation && (
            <span className={`${styles.chip} ${getRecommendationClass(candidate.recommendation)}`}>
              {candidate.recommendation}
            </span>
          )}
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
      </div>
    );
  }

  private _renderCategorySection(category: CandidateCategory, candidates: ICandidate[]): React.ReactNode {
    const { label, color } = CATEGORY_CONFIG[category];
    const isExpanded = this.state.expandedCategories.has(category);

    return (
      <div key={category} className={styles.categorySection} style={{ color }}>
        <div
          className={styles.categoryHeader}
          role="button"
          tabIndex={0}
          onClick={() => this._toggleCategory(category)}
          onKeyDown={e => {
            if (e.key === 'Enter' || e.key === ' ') {
              e.preventDefault();
              this._toggleCategory(category);
            }
          }}
        >
          <Text className={styles.categoryLabel} styles={{ root: { color } }}>{label}</Text>
          <span className={styles.categoryCount} style={{ background: color }}>
            <span className={styles.categoryCountText}>{candidates.length}</span>
          </span>
          <ActionButton
            iconProps={{ iconName: isExpanded ? 'ChevronUp' : 'ChevronDown' }}
            styles={{ root: { padding: 0, height: 'auto', minWidth: 'auto', color } }}
            ariaLabel={isExpanded ? 'Collapse' : 'Expand'}
          />
        </div>

        {isExpanded && (
          <div className={styles.categoryBody}>
            {candidates.length === 0 ? (
              <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', padding: '12px 0' } }}>
                No candidates currently in this category.
              </Text>
            ) : (
              candidates.map(c => this._renderCandidateRow(c))
            )}
          </div>
        )}
      </div>
    );
  }

  public render(): React.ReactElement {
    const { candidates, interviews, loading, error } = this.state;

    const grouped: Record<CandidateCategory, ICandidate[]> = {
      'Received': [], 'Screened': [], 'Round 1': [], 'Round 2': [],
      'HR Discussion': [], 'Final Discussion': [], 'Rejected': [],
    };
    candidates.forEach(c => { grouped[deriveCategory(c, interviews)].push(c); });

    return (
      <div className={styles.container}>
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{ root: { marginBottom: 16 } }}
        >
          <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>All Candidates</Text>
          <ActionButton
            iconProps={{ iconName: 'Refresh' }}
            text="Refresh"
            onClick={() => { this._loadData().catch(() => undefined); }}
            disabled={loading}
          />
        </Stack>

        {loading && <Spinner size={SpinnerSize.medium} label="Loading candidates…" />}

        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ error: '' })}
          >
            {error}
          </MessageBar>
        )}

        {!loading && !error && candidates.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info}>
            No candidates found across any job openings yet.
          </MessageBar>
        )}

        {!loading && !error && candidates.length > 0 && (
          <>
            <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', marginBottom: 16 } }}>
              {candidates.length} candidate{candidates.length !== 1 ? 's' : ''} across all job openings, grouped by pipeline stage. Click a category to view its candidates.
            </Text>
            {CATEGORY_ORDER.map(cat => this._renderCategorySection(cat, grouped[cat]))}
          </>
        )}
      </div>
    );
  }
}
