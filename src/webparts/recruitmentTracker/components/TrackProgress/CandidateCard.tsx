import * as React from 'react';
import {
  Stack,
  Text,
  ProgressIndicator,
  TextField,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import { SpService } from '../shared/SpService';
import { EmailService } from '../shared/EmailService';
import { ICandidate, IJobOpening, IInterview } from '../shared/models';
import { InterviewScheduler } from './InterviewScheduler';
import styles from './TrackProgress.module.scss';

interface ICandidateCardProps {
  candidate: ICandidate;
  job: IJobOpening;
  spService: SpService;
  emailService: EmailService;
}

interface ICandidateCardState {
  hrFeedback: string;
  savingFeedback: boolean;
  feedbackSaved: boolean;
  feedbackError: string;
  showScheduler: boolean;
  interviews: IInterview[];
  loadingInterviews: boolean;
}

export class CandidateCard extends React.Component<ICandidateCardProps, ICandidateCardState> {
  constructor(props: ICandidateCardProps) {
    super(props);
    this.state = {
      hrFeedback: props.candidate.hrFeedback,
      savingFeedback: false,
      feedbackSaved: false,
      feedbackError: '',
      showScheduler: false,
      interviews: [],
      loadingInterviews: true,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadInterviews();
  }

  private async _loadInterviews(): Promise<void> {
    try {
      const interviews = await this.props.spService.getInterviewsByCandidate(this.props.candidate.id);
      this.setState({ interviews, loadingInterviews: false });
    } catch {
      this.setState({ loadingInterviews: false });
    }
  }

  private _onSaveFeedback = async (): Promise<void> => {
    this.setState({ savingFeedback: true, feedbackError: '', feedbackSaved: false });
    try {
      await this.props.spService.updateHRFeedback(this.props.candidate.id, this.state.hrFeedback);
      this.setState({ savingFeedback: false, feedbackSaved: true });
      setTimeout(() => this.setState({ feedbackSaved: false }), 3000);
    } catch (err) {
      this.setState({ savingFeedback: false, feedbackError: (err as Error).message });
    }
  };

  private _onScheduled = async (): Promise<void> => {
    this.setState({ showScheduler: false });
    await this._loadInterviews();
  };

  private _renderSkillChips(skills: string, chipClass: string): React.ReactNode {
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

  private _getRecommendationClass(): string {
    switch (this.props.candidate.recommendation) {
      case 'Recommended':     return styles.badgeGreen;
      case 'Maybe':           return styles.badgeOrange;
      case 'Not Recommended': return styles.badgeRed;
      default:                return styles.badgeGrey;
    }
  }

  private _getScoreColor(score: number): string {
    if (score >= 75) return '#107c10';
    if (score >= 50) return '#f7630c';
    return '#a80000';
  }

  public render(): React.ReactElement {
    const { candidate, job, spService, emailService } = this.props;
    const {
      hrFeedback, savingFeedback, feedbackSaved, feedbackError,
      showScheduler, interviews, loadingInterviews,
    } = this.state;

    return (
      <div className={styles.candidateCard}>
        {/* Header row */}
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 8 }}>
          <Stack tokens={{ childrenGap: 2 }} styles={{ root: { flex: 1, minWidth: 180 } }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              {candidate.candidateName}
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              {candidate.email}{candidate.phone ? ` · ${candidate.phone}` : ''}
            </Text>
          </Stack>
          <span className={`${styles.badge} ${this._getRecommendationClass()}`}>
            {candidate.recommendation || 'Pending'}
          </span>
        </Stack>

        {/* Fitment score bar */}
        <div className={styles.scoreRow}>
          <Text
            variant="small"
            styles={{ root: { color: this._getScoreColor(candidate.fitmentScore), fontWeight: 600, minWidth: 44 } }}
          >
            {candidate.fitmentScore}%
          </Text>
          <div style={{ flex: 1 }}>
            <ProgressIndicator
              percentComplete={candidate.fitmentScore / 100}
              barHeight={8}
              styles={{
                itemProgress: { padding: 0 },
                progressBar: { backgroundColor: this._getScoreColor(candidate.fitmentScore) },
              }}
            />
          </div>
          {candidate.experienceMatch && (
            <span className={`${styles.chip} ${
              candidate.experienceMatch === 'meets'   ? styles.chipGreen  :
              candidate.experienceMatch === 'exceeds' ? styles.chipBlue   :
              styles.chipRed
            }`}>
              Exp: {candidate.experienceMatch}
            </span>
          )}
        </div>

        {/* Skills */}
        {candidate.matchingSkills && (
          <div className={styles.skillSection}>
            <Text variant="xSmall" styles={{ root: { color: '#605e5c', textTransform: 'uppercase', letterSpacing: '0.5px' } }}>
              Matching Skills
            </Text>
            {this._renderSkillChips(candidate.matchingSkills, styles.chipGreen)}
          </div>
        )}
        {candidate.missingSkills && (
          <div className={styles.skillSection}>
            <Text variant="xSmall" styles={{ root: { color: '#605e5c', textTransform: 'uppercase', letterSpacing: '0.5px' } }}>
              Missing Skills
            </Text>
            {this._renderSkillChips(candidate.missingSkills, styles.chipRed)}
          </div>
        )}

        {/* AI Summary */}
        {candidate.aiSummary && (
          <div className={styles.aiSummary}>
            <Text variant="xSmall" styles={{ root: { color: '#0078d4', fontWeight: 600 } }}>AI Summary</Text>
            <Text variant="small">{candidate.aiSummary}</Text>
          </div>
        )}

        {/* HR Feedback */}
        <div className={styles.feedbackSection}>
          <TextField
            label="HR Feedback"
            multiline
            rows={3}
            value={hrFeedback}
            onChange={(_, v) => this.setState({ hrFeedback: v ?? '', feedbackSaved: false })}
            disabled={savingFeedback}
            placeholder="Add notes or feedback for this candidate…"
          />
          {feedbackError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ feedbackError: '' })}>
              {feedbackError}
            </MessageBar>
          )}
          {feedbackSaved && (
            <MessageBar messageBarType={MessageBarType.success}>Feedback saved.</MessageBar>
          )}
          <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 8 } }}>
            <PrimaryButton
              text={savingFeedback ? 'Saving…' : 'Save Feedback'}
              onClick={() => { this._onSaveFeedback().catch(err => this.setState({ feedbackError: String(err) })); }}
              disabled={savingFeedback}
              iconProps={{ iconName: 'Save' }}
            />
            {savingFeedback && <Spinner size={SpinnerSize.small} />}
            <DefaultButton
              text="Schedule Interview"
              iconProps={{ iconName: 'Calendar' }}
              onClick={() => this.setState({ showScheduler: true })}
            />
            {candidate.resumeUrl && (
              <DefaultButton
                text="View Resume"
                iconProps={{ iconName: 'OpenInNewWindow' }}
                href={candidate.resumeUrl}
                target="_blank"
              />
            )}
          </Stack>
        </div>

        {/* Interview history */}
        {loadingInterviews ? (
          <Spinner size={SpinnerSize.xSmall} label="Loading interviews…" styles={{ root: { marginTop: 8 } }} />
        ) : interviews.length > 0 && (
          <div className={styles.interviewHistory}>
            <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
              Interview History
            </Text>
            {interviews.map(iv => (
              <div key={iv.id} className={styles.interviewRow}>
                <span className={styles.roundBadge}>Rd {iv.interviewRound}</span>
                <Text variant="small" styles={{ root: { flex: 1 } }}>{iv.interviewerEmail}</Text>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  {iv.scheduledDate ? new Date(iv.scheduledDate).toLocaleString('en-GB') : '—'}
                </Text>
                <span className={`${styles.chip} ${iv.feedbackStatus === 'Submitted' ? styles.chipGreen : styles.chipOrange}`}>
                  {iv.feedbackStatus}
                </span>
              </div>
            ))}
          </div>
        )}

        {/* InterviewScheduler panel (renders as Fluent UI Panel overlay) */}
        {showScheduler && (
          <InterviewScheduler
            candidate={candidate}
            job={job}
            spService={spService}
            emailService={emailService}
            onDismiss={() => this.setState({ showScheduler: false })}
            onScheduled={() => { this._onScheduled().catch(() => undefined); }}
          />
        )}
      </div>
    );
  }
}
