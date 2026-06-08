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
} from '@fluentui/react';
import { SpService } from '../shared/SpService';
import { EmailService } from '../shared/EmailService';
import { ICandidate, IJobOpening, IInterview } from '../shared/models';

interface IInterviewSchedulerProps {
  candidate: ICandidate;
  job: IJobOpening;
  spService: SpService;
  emailService: EmailService;
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
}

const INTERVIEWER_OPTIONS: IDropdownOption[] = [
  { key: 'interviewer1@company.onmicrosoft.com', text: 'Interviewer 1 (interviewer1@company.onmicrosoft.com)' },
  { key: 'jedidiah.vachan@operative.com', text: 'Jedidiah Vachan (jedidiah.vachan@operative.com)' },
  { key: 'tulsidas.rp@operative.com', text: 'Tulsidas Patil (tulsidas.rp@operative.com)' }
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
    };
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
        // Logged so a Mail.Send permission/consent issue is visible in the browser console.
        console.error('[OpATS] Failed to send interview notification emails:', emailErr);
      }

      this.setState({ submitting: false });
      onScheduled();
    } catch (err) {
      this.setState({ submitting: false, error: (err as Error).message });
    }
  };

  public render(): React.ReactElement {
    const { candidate, onDismiss } = this.props;
    const { round, interviewerEmail, scheduledDate, scheduledTime, submitting, error } = this.state;

    return (
      <Panel
        isOpen
        type={PanelType.medium}
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
        <Stack tokens={{ childrenGap: 16, padding: '16px 0 0 0' }}>
          {error && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ error: '' })}>
              {error}
            </MessageBar>
          )}

          <Text variant="medium" styles={{ root: { color: '#605e5c' } }}>
            Candidate: <strong>{candidate.candidateName}</strong>
            &nbsp;|&nbsp;Score: <strong>{candidate.fitmentScore}%</strong>
            &nbsp;|&nbsp;{candidate.recommendation}
          </Text>

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
      </Panel>
    );
  }
}
