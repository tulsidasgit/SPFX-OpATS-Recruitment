import * as React from 'react';
import {
  Dropdown,
  IDropdownOption,
  TextField,
  DatePicker,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Stack,
  Label,
  Panel,
  PanelType,
  Text,
  IconButton,
} from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { SpService } from '../shared/SpService';
import { GraphService } from '../shared/GraphService';
import { EmailService } from '../shared/EmailService';
import { AIService } from '../shared/AIService';
import { IDepartment, IJobTitle, ICurrentUser, IJobOpening } from '../shared/models';
import styles from './PostJobForm.module.scss';

interface IPostJobFormProps {
  sp: SPFI;
  graphService: GraphService;
  currentUser: ICurrentUser;
  applyBaseUrl: string;
}

interface IPostJobFormState {
  // Permission
  checkingPermission: boolean;
  isAllowedPoster: boolean;

  // Dropdowns
  departments: IDepartment[];
  jobTitles: IJobTitle[];
  loadingDepts: boolean;
  loadingTitles: boolean;

  // Form values
  selectedDeptId: number | undefined;
  selectedDeptName: string;
  selectedJobTitleId: number | undefined;
  selectedJobTitleName: string;
  mustHaveSkills: string[];
  mustHaveInput: string;
  goodToHaveSkills: string[];
  goodToHaveInput: string;
  jobLocation: string;
  jobType: IJobOpening['jobType'];
  experience: string;
  dueDate: Date | undefined;
  linkedInUrl: string;

  // Submission
  submitting: boolean;
  error: string;
  successJobRef: number | undefined;

  // JD Panel
  showJDPanel: boolean;
  generatingJD: boolean;
  jobDescription: string;
  jdError: string;
  applyUrl: string;
  linkCopied: boolean;
  pendingJobId: number | undefined;   // job created during JD generation, before JD is saved
}

const EXPERIENCE_OPTIONS: IDropdownOption[] = [
  { key: '0-1 years', text: '0–1 years' },
  { key: '1-3 years', text: '1–3 years' },
  { key: '3-5 years', text: '3–5 years' },
  { key: '5-8 years', text: '5–8 years' },
  { key: '8+ years', text: '8+ years' },
];

const LOCATION_OPTIONS: IDropdownOption[] = [
  { key: 'New York',   text: 'New York' },
  { key: 'São Paulo',  text: 'São Paulo' },
  { key: 'London',     text: 'London' },
  { key: 'Amsterdam',  text: 'Amsterdam' },
  { key: 'Craiova',    text: 'Craiova' },
  { key: 'Kyiv',       text: 'Kyiv' },
  { key: 'Jerusalem',  text: 'Jerusalem' },
  { key: 'Bangalore',  text: 'Bangalore' },
  { key: 'Sydney',     text: 'Sydney' },
];

const JOB_TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'On Site', text: 'On Site' },
  { key: 'Hybrid',  text: 'Hybrid' },
  { key: 'Remote',  text: 'Remote' },
];

export class PostJobForm extends React.Component<IPostJobFormProps, IPostJobFormState> {
  private _spService: SpService;
  private _emailService: EmailService;
  private _aiService: AIService;
  private _mustHaveInputRef = React.createRef<HTMLInputElement>();
  private _goodToHaveInputRef = React.createRef<HTMLInputElement>();

  constructor(props: IPostJobFormProps) {
    super(props);
    this._spService = new SpService(props.sp);
    this._emailService = new EmailService(props.graphService);
    this._aiService = new AIService();
    this.state = {
      checkingPermission: true,
      isAllowedPoster: false,
      departments: [],
      jobTitles: [],
      loadingDepts: false,
      loadingTitles: false,
      selectedDeptId: undefined,
      selectedDeptName: '',
      selectedJobTitleId: undefined,
      selectedJobTitleName: '',
      mustHaveSkills: [],
      mustHaveInput: '',
      goodToHaveSkills: [],
      goodToHaveInput: '',
      jobLocation: '',
      jobType: '',
      experience: '',
      dueDate: undefined,
      linkedInUrl: '',
      submitting: false,
      error: '',
      successJobRef: undefined,
      showJDPanel: false,
      generatingJD: false,
      jobDescription: '',
      jdError: '',
      applyUrl: '',
      linkCopied: false,
      pendingJobId: undefined,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._checkPermission();
  }

  // ── Permission ──────────────────────────────────────────────────────────────

  private async _checkPermission(): Promise<void> {
    try {
      const allowed = await this._spService.isAllowedPoster(this.props.currentUser.email);
      this.setState({ checkingPermission: false, isAllowedPoster: allowed });
      if (allowed) await this._loadDepartments();
    } catch (err) {
      const msg = (err as Error).message ?? '';
      const isListMissing = msg.toLowerCase().includes('does not exist') ||
                            msg.toLowerCase().includes('list') ||
                            msg.toLowerCase().includes('404') ||
                            msg.toLowerCase().includes('not found');
      this.setState({
        checkingPermission: false,
        isAllowedPoster: false,
        error: isListMissing
          ? '⚙ Setup required: The AllowedPosters SharePoint list has not been created yet. ' +
            'Please run setup/CreateLists.ps1 against your site, then reload this page.'
          : `Permission check failed: ${msg}`,
      });
    }
  }

  // ── Departments & Job Titles ────────────────────────────────────────────────

  private async _loadDepartments(): Promise<void> {
    this.setState({ loadingDepts: true });
    try {
      const departments = await this._spService.getDepartments();
      this.setState({ departments, loadingDepts: false });
    } catch (err) {
      this.setState({ loadingDepts: false, error: `Failed to load departments: ${(err as Error).message}` });
    }
  }

  private async _onDepartmentChange(_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): Promise<void> {
    if (!option) return;
    this.setState({
      selectedDeptId: option.key as number,
      selectedDeptName: option.text,
      selectedJobTitleId: undefined,
      selectedJobTitleName: '',
      jobTitles: [],
      loadingTitles: true,
    });
    try {
      const jobTitles = await this._spService.getJobTitlesByDepartment(option.key as number);
      this.setState({ jobTitles, loadingTitles: false });
    } catch (err) {
      this.setState({ loadingTitles: false, error: `Failed to load job titles: ${(err as Error).message}` });
    }
  }

  private _onJobTitleChange = (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;
    this.setState({ selectedJobTitleId: option.key as number, selectedJobTitleName: option.text });
  };

  // ── Skills tag inputs (shared logic via type param) ─────────────────────────

  private _addTag(field: 'mustHave' | 'goodToHave'): void {
    const inputKey  = field === 'mustHave' ? 'mustHaveInput'  : 'goodToHaveInput';
    const arrayKey  = field === 'mustHave' ? 'mustHaveSkills' : 'goodToHaveSkills';
    const raw = this.state[inputKey].trim().replace(/,$/, '');
    if (!raw) return;
    const existing = this.state[arrayKey] as string[];
    if (existing.indexOf(raw) === -1) {
      this.setState(prev => ({
        [arrayKey]: [...(prev[arrayKey] as string[]), raw],
        [inputKey]: '',
      } as unknown as IPostJobFormState));
    } else {
      this.setState({ [inputKey]: '' } as unknown as IPostJobFormState);
    }
  }

  private _removeTag(field: 'mustHave' | 'goodToHave', skill: string): void {
    const arrayKey = field === 'mustHave' ? 'mustHaveSkills' : 'goodToHaveSkills';
    this.setState(prev => ({
      [arrayKey]: (prev[arrayKey] as string[]).filter(s => s !== skill),
    } as unknown as IPostJobFormState));
  }

  private _onTagKeyDown = (field: 'mustHave' | 'goodToHave', e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter' || e.key === ',') {
      e.preventDefault();
      this._addTag(field);
    } else if (e.key === 'Backspace') {
      const inputKey = field === 'mustHave' ? 'mustHaveInput' : 'goodToHaveInput';
      const arrayKey = field === 'mustHave' ? 'mustHaveSkills' : 'goodToHaveSkills';
      if (this.state[inputKey] === '') {
        const arr = [...(this.state[arrayKey] as string[])];
        arr.pop();
        this.setState({ [arrayKey]: arr } as unknown as IPostJobFormState);
      }
    }
  };

  // ── Validation & submission ─────────────────────────────────────────────────

  private _isFormValid(): boolean {
    const { selectedDeptId, selectedJobTitleId, mustHaveSkills, jobLocation, jobType, experience, dueDate } = this.state;
    return (
      !!selectedDeptId &&
      !!selectedJobTitleId &&
      mustHaveSkills.length > 0 &&
      !!jobLocation &&
      !!jobType &&
      !!experience &&
      !!dueDate
    );
  }

  private _onSubmit = async (): Promise<void> => {
    if (!this._isFormValid()) return;
    this.setState({ submitting: true, error: '' });

    try {
      const {
        selectedDeptName, selectedJobTitleName,
        mustHaveSkills, goodToHaveSkills,
        jobLocation, jobType,
        experience, dueDate, linkedInUrl,
      } = this.state;

      const jobPayload: Omit<IJobOpening, 'id'> = {
        title: `${selectedJobTitleName} — ${selectedDeptName}`,
        department: selectedDeptName,
        jobTitle: selectedJobTitleName,
        requiredSkills: mustHaveSkills.join(', '),
        goodToHaveSkills: goodToHaveSkills.join(', '),
        jobLocation,
        jobType: jobType as IJobOpening['jobType'],
        experience,
        dueDate: dueDate!.toISOString(),
        status: 'Open',
        postedBy: this.props.currentUser.email,
        linkedInUrl,
      };

      const jobId = await this._spService.createJobOpening(jobPayload);

      // Build and persist the apply URL
      const applyUrl = this.props.applyBaseUrl
        ? `${this.props.applyBaseUrl.replace(/\/$/, '')}/apply?jobId=${jobId}`
        : '';
      if (applyUrl) {
        await this._spService.updateJobOpeningApplicationUrl(jobId, applyUrl);
      }

      await this._spService.ensureResumeFolder(selectedDeptName, selectedJobTitleName);

      try {
        await this._emailService.notifyHRJobPosted(
          { id: jobId, ...jobPayload },
          this.props.currentUser.displayName
        );
      } catch {
        // Email failure is non-fatal
      }

      this.setState({ submitting: false, successJobRef: jobId });
      this._resetForm();
    } catch (err) {
      this.setState({ submitting: false, error: (err as Error).message });
    }
  };

  private _resetForm(): void {
    this.setState({
      selectedDeptId: undefined,
      selectedDeptName: '',
      selectedJobTitleId: undefined,
      selectedJobTitleName: '',
      mustHaveSkills: [],
      mustHaveInput: '',
      goodToHaveSkills: [],
      goodToHaveInput: '',
      jobLocation: '',
      jobType: '',
      experience: '',
      dueDate: undefined,
      linkedInUrl: '',
      showJDPanel: false,
      jobDescription: '',
      jdError: '',
      applyUrl: '',
      linkCopied: false,
      pendingJobId: undefined,
    });
  }

  // ── JD Panel ────────────────────────────────────────────────────────────────

  private async _openJDPanel(): Promise<void> {
    this.setState({ showJDPanel: true, generatingJD: true, jdError: '', jobDescription: '' });
    try {
      const {
        selectedJobTitleName, selectedDeptName, jobLocation, jobType,
        mustHaveSkills, goodToHaveSkills, experience, dueDate, linkedInUrl,
        pendingJobId,
      } = this.state;

      // Step 1 — Create the job opening immediately so we have a real jobId for the apply URL.
      // If we already created it (Regenerate clicked), reuse the existing id.
      let jobId = pendingJobId;
      let builtApplyUrl = this.state.applyUrl;

      if (!jobId) {
        const jobPayload: Omit<IJobOpening, 'id'> = {
          title: `${selectedJobTitleName} — ${selectedDeptName}`,
          department: selectedDeptName,
          jobTitle: selectedJobTitleName,
          requiredSkills: mustHaveSkills.join(', '),
          goodToHaveSkills: goodToHaveSkills.join(', '),
          jobLocation,
          jobType: jobType as IJobOpening['jobType'],
          experience,
          dueDate: dueDate!.toISOString(),
          status: 'Open',
          postedBy: this.props.currentUser.email,
          linkedInUrl,
        };
        jobId = await this._spService.createJobOpening(jobPayload);

        // Step 2 — Build and persist the apply URL now that we have the real jobId
        builtApplyUrl = this.props.applyBaseUrl
          ? `${this.props.applyBaseUrl.replace(/\/$/, '')}/apply?jobId=${jobId}`
          : '';
        if (builtApplyUrl) {
          await this._spService.updateJobOpeningApplicationUrl(jobId, builtApplyUrl);
        }

        await this._spService.ensureResumeFolder(selectedDeptName, selectedJobTitleName);
        this.setState({ pendingJobId: jobId, applyUrl: builtApplyUrl });
      }

      // Step 3 — Generate the JD with the real URL already embedded
      const templateText = await this._spService.getJDTemplateText();
      const jd = await this._aiService.generateJobDescription(templateText, {
        jobTitle: selectedJobTitleName,
        department: selectedDeptName,
        jobLocation,
        jobType,
        mustHaveSkills: mustHaveSkills.join(', '),
        goodToHaveSkills: goodToHaveSkills.join(', '),
        experience,
        applyUrl: builtApplyUrl || undefined,
      });

      this.setState({ generatingJD: false, jobDescription: jd });
    } catch (err) {
      this.setState({ generatingJD: false, jdError: (err as Error).message });
    }
  }

  // Job is already created — just save the edited JD text back to SharePoint
  private _onPostWithJD = async (): Promise<void> => {
    const { pendingJobId, jobDescription, selectedDeptName, selectedJobTitleName } = this.state;
    if (!pendingJobId) return;

    this.setState({ submitting: true, jdError: '' });
    try {
      await this._spService.updateJobDescription(pendingJobId, jobDescription);

      try {
        const job = await this._spService.getJobOpeningById(pendingJobId);
        await this._emailService.notifyHRJobPosted(job, this.props.currentUser.displayName);
      } catch {
        // email failure is non-fatal
      }

      this.setState({ submitting: false, showJDPanel: false, successJobRef: pendingJobId });
      this._resetForm();
    } catch (err) {
      this.setState({ submitting: false, jdError: (err as Error).message });
    }
  };

  // ── Shared skill-tag renderer ───────────────────────────────────────────────

  private _renderTagInput(
    field: 'mustHave' | 'goodToHave',
    label: string,
    required: boolean
  ): React.ReactNode {
    const inputKey  = field === 'mustHave' ? 'mustHaveInput'  : 'goodToHaveInput';
    const arrayKey  = field === 'mustHave' ? 'mustHaveSkills' : 'goodToHaveSkills';
    const ref       = field === 'mustHave' ? this._mustHaveInputRef : this._goodToHaveInputRef;
    const skills    = this.state[arrayKey] as string[];
    const inputVal  = this.state[inputKey] as string;
    const tagClass  = field === 'mustHave' ? styles.tagMustHave : styles.tagGoodToHave;

    return (
      <div className={styles.fieldGroup}>
        <Label className={styles.skillsLabel} required={required}>{label}</Label>
        <div
          className={styles.skillsContainer}
          onClick={() => ref.current?.focus()}
        >
          {skills.map(skill => (
            <span key={skill} className={`${styles.tag} ${tagClass}`}>
              {skill}
              <button
                onClick={() => this._removeTag(field, skill)}
                aria-label={`Remove ${skill}`}
                disabled={this.state.submitting}
              >×</button>
            </span>
          ))}
          <input
            ref={ref}
            className={styles.tagInput}
            value={inputVal}
            onChange={e => this.setState({ [inputKey]: e.target.value } as unknown as IPostJobFormState)}
            onKeyDown={e => this._onTagKeyDown(field, e)}
            onBlur={() => { if (inputVal.trim()) this._addTag(field); }}
            placeholder={skills.length === 0 ? 'Type a skill and press Enter…' : ''}
            disabled={this.state.submitting}
          />
        </div>
      </div>
    );
  }

  // ── Render ──────────────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const {
      checkingPermission, isAllowedPoster,
      departments, jobTitles, loadingDepts, loadingTitles,
      selectedDeptId, selectedJobTitleId,
      jobLocation, jobType, experience, dueDate, linkedInUrl,
      submitting, error, successJobRef,
    } = this.state;

    if (checkingPermission) {
      return (
        <div className={styles.container}>
          <Spinner size={SpinnerSize.medium} label="Checking permissions..." />
        </div>
      );
    }

    if (!isAllowedPoster) {
      const isSetupError = error.startsWith('⚙');
      return (
        <div className={styles.container}>
          {error && (
            <MessageBar messageBarType={MessageBarType.severeWarning}>{error}</MessageBar>
          )}
          {!isSetupError && (
            <div className={styles.accessDenied}>
              <div className={styles.icon}>🔒</div>
              <h2>Access Denied</h2>
              <p>
                You do not have permission to post job openings. Only authorised users in the
                <strong> AllowedPosters</strong> list can post jobs.
              </p>
              <p>
                Your email: <strong>{this.props.currentUser.email}</strong><br />
                Contact your HR administrator at <strong>hr1@company.onmicrosoft.com</strong> to request access.
              </p>
            </div>
          )}
        </div>
      );
    }

    const deptOptions: IDropdownOption[] = departments.map(d => ({ key: d.id, text: d.title }));
    const jtOptions:   IDropdownOption[] = jobTitles.map(j => ({ key: j.id, text: j.title }));

    return (
      <>
      <div className={styles.container}>
        <p className={styles.formTitle}>Post a New Job Opening</p>
        <p className={styles.formSubtitle}>
          Posting as <strong>{this.props.currentUser.displayName}</strong>
        </p>

        {successJobRef && (
          <div className={styles.successBanner}>
            <span className={styles.successIcon}>✔</span>
            <p>
              Job posted successfully! Reference: <strong>#{successJobRef}</strong>.
              HR has been notified and a resume folder has been created.
            </p>
          </div>
        )}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ error: '' })}>
            {error}
          </MessageBar>
        )}

        <Stack tokens={{ childrenGap: 12 }}>

          {/* Row 1 — Department + Job Title */}
          <div className={styles.row}>
            <Dropdown
              label="Department"
              placeholder={loadingDepts ? 'Loading…' : 'Select department'}
              options={deptOptions}
              selectedKey={selectedDeptId}
              onChange={(e, o) => { this._onDepartmentChange(e, o).catch(err => this.setState({ error: String(err) })); }}
              disabled={loadingDepts || submitting}
              required
            />
            <Dropdown
              label="Job Title"
              placeholder={
                !selectedDeptId  ? 'Select department first' :
                loadingTitles    ? 'Loading…'                :
                                   'Select job title'
              }
              options={jtOptions}
              selectedKey={selectedJobTitleId}
              onChange={this._onJobTitleChange}
              disabled={!selectedDeptId || loadingTitles || submitting}
              required
            />
          </div>

          {/* Row 2 — Job Location + Job Type */}
          <div className={styles.row}>
            <Dropdown
              label="Job Location"
              placeholder="Select location"
              options={LOCATION_OPTIONS}
              selectedKey={jobLocation || undefined}
              onChange={(_, o) => o && this.setState({ jobLocation: o.key as string })}
              disabled={submitting}
              required
            />
            <Dropdown
              label="Job Type"
              placeholder="Select type"
              options={JOB_TYPE_OPTIONS}
              selectedKey={jobType || undefined}
              onChange={(_, o) => o && this.setState({ jobType: o.key as IJobOpening['jobType'] })}
              disabled={submitting}
              required
            />
          </div>

          {/* Row 3 — Must Have Skills */}
          {this._renderTagInput('mustHave', 'Must Have Skills', true)}

          {/* Row 4 — Good To Have Skills */}
          {this._renderTagInput('goodToHave', 'Good To Have Skills', false)}

          {/* Row 5 — Experience + Due Date */}
          <div className={styles.row}>
            <Dropdown
              label="Experience Required"
              placeholder="Select experience"
              options={EXPERIENCE_OPTIONS}
              selectedKey={experience || undefined}
              onChange={(_, o) => o && this.setState({ experience: o.key as string })}
              disabled={submitting}
              required
            />
            <DatePicker
              label="Application Due Date"
              placeholder="Select date"
              value={dueDate}
              onSelectDate={date => this.setState({ dueDate: date ?? undefined })}
              minDate={new Date()}
              disabled={submitting}
              isRequired
            />
          </div>

          {/* Row 6 — LinkedIn URL */}
          <TextField
            label="LinkedIn Job URL (optional)"
            placeholder="https://linkedin.com/jobs/..."
            value={linkedInUrl}
            onChange={(_, v) => this.setState({ linkedInUrl: v ?? '' })}
            disabled={submitting}
          />

          {/* Actions */}
          <div className={styles.actions}>
            <PrimaryButton
              text="Generate Job Description"
              iconProps={{ iconName: 'EditNote' }}
              onClick={() => { this._openJDPanel().catch(err => this.setState({ jdError: String(err) })); }}
              disabled={!this._isFormValid() || submitting}
            />
            <DefaultButton
              text={submitting ? 'Posting…' : 'Quick Post'}
              iconProps={{ iconName: 'Send' }}
              onClick={() => { this._onSubmit().catch(err => this.setState({ error: String(err) })); }}
              disabled={!this._isFormValid() || submitting}
              title="Post job immediately without generating a job description"
            />
            {submitting && <Spinner size={SpinnerSize.small} />}
            <DefaultButton
              text="Reset"
              onClick={this._resetForm.bind(this)}
              disabled={submitting}
            />
          </div>

        </Stack>
      </div>

      {/* ── JD Preview & Edit Panel ── */}
      {this._renderJDPanel()}
    </>
    );
  }

  private _renderJDPanel(): React.ReactNode {
    const { showJDPanel, generatingJD, jobDescription, jdError, submitting,
            selectedJobTitleName, selectedDeptName, jobLocation, jobType, experience,
            applyUrl, linkCopied } = this.state;
    if (!showJDPanel) return null;

    const wordCount = jobDescription.trim()
      ? jobDescription.trim().split(/\s+/).length
      : 0;

    return (
      <Panel
        isOpen={showJDPanel}
        type={PanelType.large}
        headerText="Job Description"
        onDismiss={() => this.setState({ showJDPanel: false })}
        isFooterAtBottom
        onRenderFooterContent={() => (
          <div className={styles.jdPanelFooter}>
            <DefaultButton
              text="Regenerate"
              iconProps={{ iconName: 'Refresh' }}
              onClick={() => { this._openJDPanel().catch(err => this.setState({ jdError: String(err) })); }}
              disabled={generatingJD || submitting}
            />
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text="Cancel"
                onClick={() => this.setState({ showJDPanel: false })}
                disabled={submitting}
              />
              <PrimaryButton
                text={submitting ? 'Saving…' : 'Save Job Description'}
                iconProps={{ iconName: 'Save' }}
                onClick={() => { this._onPostWithJD().catch(err => this.setState({ jdError: String(err) })); }}
                disabled={generatingJD || submitting || !jobDescription.trim()}
              />
              {submitting && <Spinner size={SpinnerSize.small} />}
            </Stack>
          </div>
        )}
      >
        {/* Job meta chips */}
        <div className={styles.jdMeta}>
          {selectedJobTitleName && <span>💼 {selectedJobTitleName}</span>}
          {selectedDeptName    && <span>🏢 {selectedDeptName}</span>}
          {jobLocation         && <span>📍 {jobLocation}</span>}
          {jobType             && <span>🕐 {jobType}</span>}
          {experience          && <span>⏳ {experience}</span>}
          {applyUrl && (
            <IconButton
              iconProps={{ iconName: linkCopied ? 'CheckMark' : 'Copy' }}
              title={linkCopied ? 'Link copied!' : 'Copy apply link'}
              styles={{
                root: {
                  background: linkCopied ? '#dff6dd' : '#deecf9',
                  borderRadius: 12,
                  height: 26,
                  padding: '0 10px',
                  color: linkCopied ? '#107c10' : '#0078d4',
                },
              }}
              onClick={() => {
                navigator.clipboard.writeText(applyUrl).catch(() => undefined);
                this.setState({ linkCopied: true });
                setTimeout(() => this.setState({ linkCopied: false }), 2500);
              }}
            />
          )}
        </div>
        {applyUrl && (
          <Text
            variant="small"
            styles={{ root: { color: '#605e5c', marginBottom: 12, wordBreak: 'break-all' } }}
          >
            Apply link: <strong>{applyUrl}</strong>
          </Text>
        )}

        {jdError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ jdError: '' })}
            styles={{ root: { marginBottom: 12 } }}
          >
            {jdError}
          </MessageBar>
        )}

        {generatingJD ? (
          <div className={styles.jdGenerating}>
            <Spinner size={SpinnerSize.large} />
            <Text variant="medium">Generating job description…</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Reading your template and crafting content with AI
            </Text>
          </div>
        ) : (
          <>
            <TextField
              multiline
              rows={24}
              value={jobDescription}
              onChange={(_, v) => this.setState({ jobDescription: v ?? '' })}
              disabled={submitting}
              styles={{
                field: {
                  fontFamily: "'Segoe UI', sans-serif",
                  fontSize: 14,
                  lineHeight: '1.6',
                  resize: 'vertical',
                },
                fieldGroup: { minHeight: 480 },
              }}
              placeholder="Job description will appear here…"
            />
            <div className={styles.jdWordCount}>
              {wordCount} words
            </div>
          </>
        )}
      </Panel>
    );
  }
}
