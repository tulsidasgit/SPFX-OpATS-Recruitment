import * as React from 'react';
import {
  Panel,
  PanelType,
  PrimaryButton,
  DefaultButton,
  TextField,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Stack,
  Text,
  Icon,
} from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { SpService } from '../shared/SpService';
import { GraphService } from '../shared/GraphService';
import { EmailService } from '../shared/EmailService';
import { IJobOpening, ICurrentUser } from '../shared/models';
import styles from './OngoingPositions.module.scss';

interface IOngoingPositionsProps {
  sp: SPFI;
  graphService: GraphService;
  currentUser: ICurrentUser;
}

interface IOngoingPositionsState {
  jobs: IJobOpening[];
  loading: boolean;
  error: string;

  applyPanelOpen: boolean;
  applyingToJob: IJobOpening | undefined;

  // Candidate details
  candidateName: string;
  email: string;
  phone: string;
  resumeFile: File | undefined;

  // Referrer details (pre-filled from logged-in user)
  referrerEmployeeId: string;
  referrerDesignation: string;

  submitting: boolean;
  submitError: string;
  successJobTitle: string;
}

export class OngoingPositions extends React.Component<IOngoingPositionsProps, IOngoingPositionsState> {
  private _spService: SpService;
  private _emailService: EmailService;
  private _fileInputRef = React.createRef<HTMLInputElement>();

  constructor(props: IOngoingPositionsProps) {
    super(props);
    this._spService = new SpService(props.sp);
    this._emailService = new EmailService(props.graphService);
    this.state = {
      jobs: [],
      loading: true,
      error: '',
      applyPanelOpen: false,
      applyingToJob: undefined,
      candidateName: '',
      email: '',
      phone: '',
      resumeFile: undefined,
      referrerEmployeeId: '',
      referrerDesignation: '',
      submitting: false,
      submitError: '',
      successJobTitle: '',
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadJobs();
  }

  // ── Data loading ────────────────────────────────────────────────────────────

  private async _loadJobs(): Promise<void> {
    this.setState({ loading: true, error: '' });
    try {
      const jobs = await this._spService.getOpenJobOpenings();
      this.setState({ jobs, loading: false });
    } catch (err) {
      this.setState({ loading: false, error: (err as Error).message });
    }
  }

  // ── Panel ───────────────────────────────────────────────────────────────────

  private _openApplyPanel(job: IJobOpening): void {
    this.setState({
      applyPanelOpen: true,
      applyingToJob: job,
      candidateName: '',
      email: '',
      phone: '',
      resumeFile: undefined,
      referrerEmployeeId: '',
      referrerDesignation: '',
      submitError: '',
    });
  }

  private _closePanel(): void {
    if (this.state.submitting) return;
    this.setState({ applyPanelOpen: false, applyingToJob: undefined, submitError: '' });
  }

  // ── Validation ──────────────────────────────────────────────────────────────

  private _isFormValid(): boolean {
    const { candidateName, email, resumeFile, referrerEmployeeId, referrerDesignation } = this.state;
    return (
      !!candidateName.trim() &&
      !!email.trim() &&
      !!resumeFile &&
      !!referrerEmployeeId.trim() &&
      !!referrerDesignation.trim()
    );
  }

  // ── Submit ──────────────────────────────────────────────────────────────────

  private _onSubmit = async (): Promise<void> => {
    const {
      applyingToJob, candidateName, email, phone, resumeFile,
      referrerEmployeeId, referrerDesignation,
    } = this.state;
    if (!applyingToJob || !resumeFile || !this._isFormValid()) return;

    const { currentUser } = this.props;
    this.setState({ submitting: true, submitError: '' });

    try {
      // Upload to Referred sub-folder
      const resumeUrl = await this._spService.uploadResume(
        applyingToJob.department,
        applyingToJob.jobTitle,
        resumeFile,
        candidateName,
        'Referred'
      );

      // Create candidate record with referral metadata
      await this._spService.createCandidate({
        jobOpeningId: applyingToJob.id,
        candidateName: candidateName.trim(),
        email: email.trim(),
        phone: phone.trim(),
        resumeUrl,
        fitmentScore: 0,
        matchingSkills: '',
        missingSkills: '',
        aiSummary: '',
        hrFeedback: '',
        recommendation: '',
        experienceMatch: '',
        referredBy: currentUser.displayName,
        referrerEmail: currentUser.email,
        referrerEmployeeId: referrerEmployeeId.trim(),
        referrerDesignation: referrerDesignation.trim(),
      });

      // Notify HR with referral details
      try {
        await this._emailService.notifyHRNewApplication(
          applyingToJob,
          candidateName.trim(),
          email.trim(),
          phone.trim(),
          resumeUrl,
          {
            referredBy: currentUser.displayName,
            referrerEmail: currentUser.email,
            employeeId: referrerEmployeeId.trim(),
            designation: referrerDesignation.trim(),
          }
        );
      } catch {
        // email failure is non-fatal
      }

      this.setState({
        submitting: false,
        successJobTitle: applyingToJob.title,
        applyPanelOpen: false,
        applyingToJob: undefined,
      });
    } catch (err) {
      this.setState({ submitting: false, submitError: (err as Error).message });
    }
  };

  // ── Helpers ─────────────────────────────────────────────────────────────────

  private _getDaysRemaining(dueDate: string): number {
    if (!dueDate) return 0;
    const due = new Date(dueDate);
    const now = new Date();
    return Math.ceil((due.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
  }

  // ── Job card ─────────────────────────────────────────────────────────────────

  private _renderJobCard(job: IJobOpening): React.ReactNode {
    const days = this._getDaysRemaining(job.dueDate);
    const daysText =
      days > 0  ? `${days} day${days === 1 ? '' : 's'} left` :
      days === 0 ? 'Closes today' : 'Expired';
    const daysClass =
      days > 7  ? styles.daysGood    :
      days >= 0 ? styles.daysWarning : styles.daysClosed;

    const skills = job.requiredSkills
      ? job.requiredSkills.split(',').map(s => s.trim()).filter(Boolean)
      : [];

    return (
      <div key={job.id} className={styles.jobCard}>
        <div className={styles.cardHeader}>
          <div>
            <div className={styles.jobTitle}>{job.title}</div>
            <div className={styles.jobMeta}>
              {job.jobLocation && (
                <span className={styles.metaChip}>
                  <Icon iconName="MapPin" />
                  {job.jobLocation}
                </span>
              )}
              {job.jobType && (
                <span className={styles.metaChip}>
                  <Icon iconName="Clock" />
                  {job.jobType}
                </span>
              )}
              {job.experience && (
                <span className={styles.metaChip}>
                  <Icon iconName="AccountActivity" />
                  {job.experience}
                </span>
              )}
            </div>
          </div>
          <span
            className={`${styles.statusBadge} ${
              job.status === 'Open' ? styles.badgeOpen : styles.badgeInProgress
            }`}
          >
            {job.status}
          </span>
        </div>

        {skills.length > 0 && (
          <div className={styles.skillsRow}>
            <span className={styles.skillsRowLabel}>Skills:</span>
            <div className={styles.skillChips}>
              {skills.slice(0, 5).map(s => (
                <span key={s} className={styles.skillChip}>{s}</span>
              ))}
              {skills.length > 5 && (
                <span className={styles.skillChipMore}>+{skills.length - 5} more</span>
              )}
            </div>
          </div>
        )}

        {job.jobDescription && (
          <div className={styles.jobDescSnippet}>{job.jobDescription}</div>
        )}

        <div className={styles.cardFooter}>
          <span className={`${styles.daysLabel} ${daysClass}`}>{daysText}</span>
          <PrimaryButton
            text="Refer / Apply"
            iconProps={{ iconName: 'People' }}
            onClick={() => this._openApplyPanel(job)}
            styles={{ root: { background: '#fd800b', borderColor: '#fd800b' } }}
          />
        </div>
      </div>
    );
  }

  // ── Apply panel ───────────────────────────────────────────────────────────────

  private _renderApplyPanel(): React.ReactNode {
    const {
      applyPanelOpen, applyingToJob,
      candidateName, email, phone, resumeFile,
      referrerEmployeeId, referrerDesignation,
      submitting, submitError,
    } = this.state;
    const { currentUser } = this.props;

    return (
      <Panel
        isOpen={applyPanelOpen}
        type={PanelType.medium}
        headerText={`Refer for: ${applyingToJob?.title ?? ''}`}
        onDismiss={() => this._closePanel()}
        isFooterAtBottom
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="end">
            <DefaultButton text="Cancel" onClick={() => this._closePanel()} disabled={submitting} />
            <PrimaryButton
              text={submitting ? 'Submitting…' : 'Submit Referral'}
              iconProps={{ iconName: 'Send' }}
              onClick={() => {
                this._onSubmit().catch(err =>
                  this.setState({ submitError: String(err), submitting: false })
                );
              }}
              disabled={!this._isFormValid() || submitting}
              styles={{ root: { background: '#fd800b', borderColor: '#fd800b' } }}
            />
            {submitting && <Spinner size={SpinnerSize.small} />}
          </Stack>
        )}
      >
        {applyingToJob && (
          <Stack tokens={{ childrenGap: 0, padding: '16px 0' }}>

            {/* Job summary */}
            <div className={styles.panelJobSummary}>
              <div className={styles.panelJobTitle}>{applyingToJob.title}</div>
              <div className={styles.panelJobMeta}>
                {[applyingToJob.department, applyingToJob.jobLocation,
                  applyingToJob.jobType, applyingToJob.experience]
                  .filter(Boolean).join(' • ')}
              </div>
            </div>

            {submitError && (
              <MessageBar
                messageBarType={MessageBarType.error}
                onDismiss={() => this.setState({ submitError: '' })}
                styles={{ root: { marginTop: 12 } }}
              >
                {submitError}
              </MessageBar>
            )}

            {/* ── Candidate details ── */}
            <div className={styles.sectionHeading}>Candidate Details</div>

            <Stack tokens={{ childrenGap: 12 }}>
              <TextField
                label="Full Name"
                placeholder="e.g. Priya Sharma"
                value={candidateName}
                onChange={(_, v) => this.setState({ candidateName: v ?? '' })}
                disabled={submitting}
                required
              />
              <TextField
                label="Email Address"
                placeholder="e.g. priya@example.com"
                value={email}
                onChange={(_, v) => this.setState({ email: v ?? '' })}
                disabled={submitting}
                required
                type="email"
              />
              <TextField
                label="Phone Number"
                placeholder="e.g. +91 98765 43210"
                value={phone}
                onChange={(_, v) => this.setState({ phone: v ?? '' })}
                disabled={submitting}
              />

              {/* Resume upload */}
              <div>
                <label className={styles.fileLabel}>
                  Resume <span className={styles.required}>*</span>
                </label>
                <div
                  className={`${styles.dropZone} ${resumeFile ? styles.dropZoneHasFile : ''}`}
                  onClick={() => !submitting && this._fileInputRef.current?.click()}
                  onDragOver={e => e.preventDefault()}
                  onDrop={e => {
                    e.preventDefault();
                    if (submitting) return;
                    const file = e.dataTransfer.files[0];
                    if (file) this.setState({ resumeFile: file });
                  }}
                >
                  <input
                    ref={this._fileInputRef}
                    type="file"
                    accept=".pdf,.doc,.docx"
                    style={{ display: 'none' }}
                    onChange={e => {
                      const file = e.target.files?.[0];
                      if (file) this.setState({ resumeFile: file });
                    }}
                  />
                  {resumeFile ? (
                    <div className={styles.fileSelected}>
                      <Icon iconName="TextDocument" />
                      <span>{resumeFile.name}</span>
                      <button
                        className={styles.removeFile}
                        onClick={e => {
                          e.stopPropagation();
                          this.setState({ resumeFile: undefined });
                          if (this._fileInputRef.current) this._fileInputRef.current.value = '';
                        }}
                        aria-label="Remove file"
                        disabled={submitting}
                      >×</button>
                    </div>
                  ) : (
                    <div className={styles.dropPrompt}>
                      <Icon iconName="Upload" className={styles.uploadIcon} />
                      <span>Click to browse or drag &amp; drop</span>
                      <span className={styles.fileHint}>PDF, DOC, DOCX · Max 10 MB</span>
                    </div>
                  )}
                </div>
                <p className={styles.folderNote}>
                  Resume will be saved in the <strong>Referred</strong> folder for this position.
                </p>
              </div>
            </Stack>

            {/* ── Referrer details ── */}
            <div className={styles.sectionDivider} />
            <div className={styles.sectionHeading}>
              <Icon iconName="Contact" /> Your Details (Referrer)
            </div>

            <Stack tokens={{ childrenGap: 12 }}>
              {/* Read-only: name + email from AD */}
              <div className={styles.referrerReadonlyRow}>
                <div className={styles.referrerChip}>
                  <span className={styles.referrerChipLabel}>Name</span>
                  <span className={styles.referrerChipValue}>{currentUser.displayName}</span>
                </div>
                <div className={styles.referrerChip}>
                  <span className={styles.referrerChipLabel}>Email</span>
                  <span className={styles.referrerChipValue}>{currentUser.email}</span>
                </div>
              </div>

              <TextField
                label="Employee ID"
                placeholder="e.g. EMP-1042"
                value={referrerEmployeeId}
                onChange={(_, v) => this.setState({ referrerEmployeeId: v ?? '' })}
                disabled={submitting}
                required
              />
              <TextField
                label="Your Designation"
                placeholder="e.g. Senior Software Engineer"
                value={referrerDesignation}
                onChange={(_, v) => this.setState({ referrerDesignation: v ?? '' })}
                disabled={submitting}
                required
              />
            </Stack>

          </Stack>
        )}
      </Panel>
    );
  }

  // ── Root render ───────────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const { jobs, loading, error, successJobTitle } = this.state;

    return (
      <div className={styles.container}>

        <div className={styles.pageHeader}>
          <Text variant="xLarge" className={styles.pageTitle}>Ongoing Positions</Text>
          <Text variant="medium" className={styles.pageSubtitle}>
            Browse open roles and refer a colleague or submit your own application
          </Text>
        </div>

        {successJobTitle && (
          <MessageBar
            messageBarType={MessageBarType.success}
            onDismiss={() => this.setState({ successJobTitle: '' })}
            styles={{ root: { marginBottom: 16 } }}
          >
            Referral submitted for <strong>{successJobTitle}</strong>! HR has been notified with your details.
          </MessageBar>
        )}

        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ error: '' })}
            styles={{ root: { marginBottom: 16 } }}
          >
            {error}
          </MessageBar>
        )}

        {loading ? (
          <div className={styles.loadingState}>
            <Spinner size={SpinnerSize.large} label="Loading open positions…" />
          </div>
        ) : jobs.length === 0 ? (
          <div className={styles.emptyState}>
            <Icon iconName="Briefcase" className={styles.emptyIcon} />
            <Text variant="large">No open positions at this time</Text>
            <Text variant="medium" styles={{ root: { color: '#605e5c' } }}>
              Check back soon — new opportunities are posted regularly.
            </Text>
          </div>
        ) : (
          <div className={styles.jobGrid}>
            {jobs.map(job => this._renderJobCard(job))}
          </div>
        )}

        {this._renderApplyPanel()}

      </div>
    );
  }
}
