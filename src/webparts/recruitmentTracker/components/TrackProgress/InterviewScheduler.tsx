import * as React from 'react';
import {
  Panel,
  PanelType,
  Dropdown,
  IDropdownOption,
  DatePicker,
  TextField,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Stack,
  Text,
  Separator,
  ProgressIndicator,
} from '@fluentui/react';
import { SpService } from '../shared/SpService';
import { EmailService } from '../shared/EmailService';
import { AIService } from '../shared/AIService';
import { ICandidate, IJobOpening, IInterview } from '../shared/models';

interface IInterviewSchedulerProps {
  candidate: ICandidate;
  job: IJobOpening;
  spService: SpService;
  emailService: EmailService;
  aiService: AIService;
  onDismiss: () => void;
  onScheduled: () => void;
}

interface IInterviewSchedulerState {
  round: IInterview['interviewRound'];
  interviewerEmail: string;
  scheduledDate: Date | undefined;
  scheduledTime: string;
  submitting: boolean;
  error: string;
  screeningQuestions: string[];
  loadingQuestions: boolean;
  questionsError: string;
  overviewExpanded: boolean;
}

const INTERVIEWER_OPTIONS: IDropdownOption[] = [
  { key: 'interviewer1@company.onmicrosoft.com', text: 'Interviewer 1 (interviewer1@company.onmicrosoft.com)' },
  { key: 'jedidiah.vachan@operative.com', text: 'Jedidiah Vachan (jedidiah.vachan@operative.com)' },
  { key: 'tulsidas.rp@operative.com', text: 'Tulsidas Patil (tulsidas.rp@operative.com)' },
];

const ROUND_OPTIONS: IDropdownOption[] = [
  { key: '1', text: 'Round 1' },
  { key: '2', text: 'Round 2' },
  { key: '3', text: 'Round 3' },
];

export class InterviewScheduler extends React.Component<IInterviewSchedulerProps, IInterviewSchedulerState> {
  constructor(props: IInterviewSchedulerProps) {
    super(props);
    this.state = {
      round: '1',
      interviewerEmail: '',
      scheduledDate: undefined,
      scheduledTime: '10:00',
      submitting: false,
      error: '',
      screeningQuestions: [],
      loadingQuestions: true,
      questionsError: '',
      overviewExpanded: true,
    };
  }

  public async componentDidMount(): Promise<void> {
    const { aiService, candidate, job } = this.props;
    try {
      const questions = await aiService.generateScreeningQuestions(candidate, job);
      this.setState({ screeningQuestions: questions, loadingQuestions: false });
    } catch (err) {
      this.setState({ loadingQuestions: false, questionsError: (err as Error).message });
    }
  }

  private _isValid(): boolean {
    const { round, interviewerEmail, scheduledDate, scheduledTime } = this.state;
    return !!round && !!interviewerEmail && !!scheduledDate && !!scheduledTime;
  }

  private _buildScheduledDateTime(): string {
    const { scheduledDate, scheduledTime } = this.state;
    if (!scheduledDate) return '';
    const parts = scheduledTime.split(':');
    const hours = parseInt(parts[0] ?? '10', 10);
    const minutes = parseInt(parts[1] ?? '0', 10);
    const dt = new Date(scheduledDate);
    dt.setHours(hours, minutes, 0, 0);
    return dt.toISOString();
  }

  private _onSubmit = async (): Promise<void> => {
    if (!this._isValid()) return;
    this.setState({ submitting: true, error: '' });

    const { candidate, job, spService, emailService, onScheduled } = this.props;
    const { round, interviewerEmail } = this.state;

    try {
      const scheduledDateTime = this._buildScheduledDateTime();

      const interviewId = await spService.createInterview({
        candidateId: candidate.id,
        jobOpeningId: job.id,
        interviewRound: round,
        interviewerEmail,
        scheduledDate: scheduledDateTime,
        feedbackStatus: 'Pending',
        feedback: '',
        hrNotes: '',
      });

      try {
        const interview: IInterview = {
          id: interviewId,
          candidateId: candidate.id,
          jobOpeningId: job.id,
          interviewRound: round,
          interviewerEmail,
          scheduledDate: scheduledDateTime,
          feedbackStatus: 'Pending',
          feedback: '',
          hrNotes: '',
        };
        await Promise.all([
          emailService.notifyInterviewerScheduled(interview, candidate, job),
          emailService.notifyCandidateInterviewScheduled(interview, candidate, job),
        ]);
      } catch (emailErr) {
        // Email failure is non-fatal — interview is already created.
        console.error('[OpRea] Failed to send interview notification emails:', emailErr);
      }

      this.setState({ submitting: false });
      onScheduled();
    } catch (err) {
      this.setState({ submitting: false, error: (err as Error).message });
    }
  };

  private _renderChips(skills: string, bg: string, color: string): React.ReactNode {
    const list = skills.split(',').map(s => s.trim()).filter(Boolean);
    if (list.length === 0) {
      return <Text variant="small" styles={{ root: { color: '#605e5c' } }}>None identified</Text>;
    }
    return (
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginTop: 4 }}>
        {list.map(s => (
          <span key={s} style={{ background: bg, color, padding: '2px 8px', borderRadius: 10, fontSize: 12, fontWeight: 500 }}>
            {s}
          </span>
        ))}
      </div>
    );
  }

  private _getScoreColor(score: number): string {
    if (score >= 75) return '#107c10';
    if (score >= 50) return '#f7630c';
    return '#a80000';
  }

  private _renderCandidateOverview(): React.ReactNode {
    const { candidate } = this.props;
    const { overviewExpanded } = this.state;
    const scoreColor = this._getScoreColor(candidate.fitmentScore);
    const recBg   = candidate.recommendation === 'Recommended' ? '#dff6dd' : candidate.recommendation === 'Maybe' ? '#fff4ce' : '#fde7e9';
    const recColor = candidate.recommendation === 'Recommended' ? '#107c10' : candidate.recommendation === 'Maybe' ? '#8a8000' : '#a80000';
    const expBg   = candidate.experienceMatch === 'meets' ? '#dff6dd' : candidate.experienceMatch === 'exceeds' ? '#deecf9' : '#fde7e9';
    const expColor = candidate.experienceMatch === 'meets' ? '#107c10' : candidate.experienceMatch === 'exceeds' ? '#0078d4' : '#a80000';

    return (
      <div style={{ marginBottom: 4 }}>
        <div
          onClick={() => this.setState({ overviewExpanded: !overviewExpanded })}
          style={{ display: 'flex', alignItems: 'center', cursor: 'pointer', padding: '4px 0 8px', userSelect: 'none' }}
        >
          <span style={{ fontSize: 12, marginRight: 8, color: '#605e5c' }}>{overviewExpanded ? '▼' : '▶'}</span>
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>Candidate Overview</Text>
        </div>

        {overviewExpanded && (
          <div style={{ border: '1px solid #edebe9', borderRadius: 6, padding: '14px 16px', background: '#faf9f8' }}>

            {/* Name + Resume button */}
            <Stack horizontal horizontalAlign="space-between" verticalAlign="start" tokens={{ childrenGap: 8 }} wrap>
              <div>
                <Text variant="large" styles={{ root: { fontWeight: 600 } }}>{candidate.candidateName}</Text>
                <div style={{ marginTop: 2 }}>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    {candidate.email}{candidate.phone ? ` · ${candidate.phone}` : ''}
                  </Text>
                </div>
              </div>
              {candidate.resumeUrl && (
                <a
                  href={candidate.resumeUrl}
                  target="_blank"
                  rel="noreferrer"
                  style={{
                    display: 'inline-flex', alignItems: 'center', gap: 6,
                    background: '#0078d4', color: '#fff', padding: '6px 14px',
                    borderRadius: 4, textDecoration: 'none', fontSize: 13, fontWeight: 600,
                    whiteSpace: 'nowrap',
                  }}
                >
                  ↓ View Resume
                </a>
              )}
            </Stack>

            {/* Score + badges */}
            <div style={{ marginTop: 12, display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 8 }}>
              <span style={{ background: scoreColor, color: '#fff', padding: '3px 12px', borderRadius: 12, fontSize: 13, fontWeight: 700 }}>
                Score: {candidate.fitmentScore}%
              </span>
              {candidate.recommendation && (
                <span style={{ background: recBg, color: recColor, padding: '3px 10px', borderRadius: 12, fontSize: 12, fontWeight: 600 }}>
                  {candidate.recommendation}
                </span>
              )}
              {candidate.experienceMatch && (
                <span style={{ background: expBg, color: expColor, padding: '3px 10px', borderRadius: 12, fontSize: 12, fontWeight: 600 }}>
                  Exp: {candidate.experienceMatch}
                </span>
              )}
            </div>

            {/* Score bar */}
            <div style={{ marginTop: 6 }}>
              <ProgressIndicator
                percentComplete={candidate.fitmentScore / 100}
                barHeight={6}
                styles={{ itemProgress: { padding: 0 }, progressBar: { backgroundColor: scoreColor } }}
              />
            </div>

            {/* Skills */}
            {candidate.matchingSkills && (
              <div style={{ marginTop: 12 }}>
                <Text variant="xSmall" styles={{ root: { color: '#605e5c', textTransform: 'uppercase', letterSpacing: '0.5px' } }}>
                  Matching Skills
                </Text>
                {this._renderChips(candidate.matchingSkills, '#dff6dd', '#107c10')}
              </div>
            )}
            {candidate.missingSkills && (
              <div style={{ marginTop: 8 }}>
                <Text variant="xSmall" styles={{ root: { color: '#605e5c', textTransform: 'uppercase', letterSpacing: '0.5px' } }}>
                  Missing Skills
                </Text>
                {this._renderChips(candidate.missingSkills, '#fde7e9', '#a80000')}
              </div>
            )}

            {/* AI Screening Summary */}
            {candidate.aiSummary && (
              <div style={{ marginTop: 12, background: '#deecf9', borderLeft: '3px solid #0078d4', padding: '8px 12px', borderRadius: '0 4px 4px 0' }}>
                <Text variant="xSmall" styles={{ root: { color: '#0078d4', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' } }}>
                  AI Screening Summary
                </Text>
                <Text variant="small" styles={{ root: { display: 'block', marginTop: 4 } }}>{candidate.aiSummary}</Text>
              </div>
            )}

            {/* HR Feedback */}
            {candidate.hrFeedback && (
              <div style={{ marginTop: 10, background: '#fff4ce', borderLeft: '3px solid #f7630c', padding: '8px 12px', borderRadius: '0 4px 4px 0' }}>
                <Text variant="xSmall" styles={{ root: { color: '#ca5010', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' } }}>
                  HR Feedback
                </Text>
                <Text variant="small" styles={{ root: { display: 'block', marginTop: 4 } }}>{candidate.hrFeedback}</Text>
              </div>
            )}
          </div>
        )}
      </div>
    );
  }

  private _renderScreeningQuestions(): React.ReactNode {
    const { screeningQuestions, loadingQuestions, questionsError } = this.state;

    return (
      <div>
        <div style={{ marginBottom: 4 }}>
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>AI Screening Questions</Text>
        </div>
        <Text variant="xSmall" styles={{ root: { color: '#605e5c', display: 'block', marginBottom: 8 } }}>
          Generated based on this candidate&apos;s profile and the job requirements
        </Text>

        {loadingQuestions && (
          <Spinner size={SpinnerSize.small} label="Generating tailored questions…" labelPosition="right" />
        )}

        {questionsError && !loadingQuestions && (
          <MessageBar
            messageBarType={MessageBarType.warning}
            onDismiss={() => this.setState({ questionsError: '' })}
          >
            Could not generate questions: {questionsError}
          </MessageBar>
        )}

        {!loadingQuestions && screeningQuestions.length > 0 && (
          <ol style={{ margin: '4px 0 0 0', paddingLeft: 20 }}>
            {screeningQuestions.map((q, i) => (
              <li key={i} style={{ marginBottom: 10, fontSize: 13, color: '#323130', lineHeight: 1.6 }}>
                {q}
              </li>
            ))}
          </ol>
        )}
      </div>
    );
  }

  public render(): React.ReactElement {
    const { candidate, onDismiss } = this.props;
    const { round, interviewerEmail, scheduledDate, scheduledTime, submitting, error } = this.state;

    return (
      <Panel
        isOpen
        type={PanelType.large}
        headerText={`Schedule Interview — ${candidate.candidateName}`}
        onDismiss={onDismiss}
        isFooterAtBottom
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text={submitting ? 'Scheduling…' : 'Schedule Interview'}
              onClick={() => { this._onSubmit().catch(err => this.setState({ error: String(err) })); }}
              disabled={!this._isValid() || submitting}
            />
            {submitting && <Spinner size={SpinnerSize.small} />}
            <DefaultButton text="Cancel" onClick={onDismiss} disabled={submitting} />
          </Stack>
        )}
      >
        <Stack tokens={{ childrenGap: 0, padding: '16px 0 0 0' }}>
          {error && (
            <MessageBar
              messageBarType={MessageBarType.error}
              onDismiss={() => this.setState({ error: '' })}
              styles={{ root: { marginBottom: 16 } }}
            >
              {error}
            </MessageBar>
          )}

          {/* Candidate overview: resume, AI analysis, HR feedback */}
          {this._renderCandidateOverview()}

          <Separator />

          {/* Interview scheduling form */}
          <Stack tokens={{ childrenGap: 12, padding: '4px 0 4px' }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>Schedule Interview</Text>

            <Dropdown
              label="Interview Round"
              selectedKey={round}
              options={ROUND_OPTIONS}
              onChange={(_, o) => o && this.setState({ round: o.key as IInterview['interviewRound'] })}
              disabled={submitting}
              required
            />

            <Dropdown
              label="Interviewer"
              placeholder="Select interviewer"
              selectedKey={interviewerEmail || undefined}
              options={INTERVIEWER_OPTIONS}
              onChange={(_, o) => o && this.setState({ interviewerEmail: o.key as string })}
              disabled={submitting}
              required
            />

            <DatePicker
              label="Interview Date"
              placeholder="Select date"
              value={scheduledDate}
              onSelectDate={date => this.setState({ scheduledDate: date ?? undefined })}
              minDate={new Date()}
              disabled={submitting}
              isRequired
            />

            <TextField
              label="Time (24h, e.g. 14:30)"
              value={scheduledTime}
              onChange={(_, v) => this.setState({ scheduledTime: v ?? '10:00' })}
              placeholder="10:00"
              disabled={submitting}
              required
            />
          </Stack>

          <Separator />

          {/* AI-generated screening questions */}
          {this._renderScreeningQuestions()}
        </Stack>
      </Panel>
    );
  }
}
