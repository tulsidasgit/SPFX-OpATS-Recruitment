import React, { useEffect, useState, useRef } from 'react';
import {
  fetchJobDetails,
  submitApplication,
  fileToBase64,
  IJobDetails,
} from './services/apiService';

interface IFormState {
  name: string;
  email: string;
  phone: string;
  currentCtc: string;
  expectedCtc: string;
  noticePeriod: string;
  gender: string;
  reasonForLeaving: string;
  workCultureExpectation: string;
  linkedInUrl: string;
}

const NOTICE_OPTIONS = ['Immediate', '15 days', '30 days', '60 days', '90 days'];
const GENDER_OPTIONS = ['Male', 'Female', 'Non-binary', 'Prefer not to say'];

const EMPTY_FORM: IFormState = {
  name: '',
  email: '',
  phone: '',
  currentCtc: '',
  expectedCtc: '',
  noticePeriod: '',
  gender: '',
  reasonForLeaving: '',
  workCultureExpectation: '',
  linkedInUrl: '',
};

interface IProps { jobId: string; }

type PageState = 'loading' | 'error' | 'form' | 'submitting' | 'success';

export default function Apply({ jobId }: IProps): React.ReactElement {
  const [pageState, setPageState] = useState<PageState>('loading');
  const [job, setJob] = useState<IJobDetails | null>(null);
  const [loadError, setLoadError] = useState('');

  const [form, setForm] = useState<IFormState>(EMPTY_FORM);
  const [resumeFile, setResumeFile] = useState<File | null>(null);
  const [resumeError, setResumeError] = useState('');
  const [fieldErrors, setFieldErrors] = useState<Partial<IFormState & { resume: string }>>({});
  const [submitError, setSubmitError] = useState('');
  const [isDragging, setIsDragging] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── Load job details ────────────────────────────────────────────────────────
  useEffect(() => {
    fetchJobDetails(jobId)
      .then(data => { setJob(data); setPageState('form'); })
      .catch(err => { setLoadError((err as Error).message); setPageState('error'); });
  }, [jobId]);

  // ── Field helpers ───────────────────────────────────────────────────────────
  const set = (field: keyof IFormState) =>
    (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
      setForm(prev => ({ ...prev, [field]: e.target.value }));
      setFieldErrors(prev => ({ ...prev, [field]: undefined }));
    };

  // ── Resume handling ─────────────────────────────────────────────────────────
  const handleFile = (file: File | null): void => {
    if (!file) return;
    const ext = file.name.split('.').pop()?.toLowerCase() ?? '';
    if (!['pdf', 'docx', 'doc'].includes(ext)) {
      setResumeError('Only PDF or DOCX files are accepted.');
      return;
    }
    if (file.size > 5 * 1024 * 1024) {
      setResumeError('File must be under 5 MB.');
      return;
    }
    setResumeFile(file);
    setResumeError('');
    setFieldErrors(prev => ({ ...prev, resume: undefined }));
  };

  const onFileInputChange = (e: React.ChangeEvent<HTMLInputElement>): void =>
    handleFile(e.target.files?.[0] ?? null);

  const onDrop = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    setIsDragging(false);
    handleFile(e.dataTransfer.files?.[0] ?? null);
  };

  // ── Validation ──────────────────────────────────────────────────────────────
  const validate = (): boolean => {
    const errors: Partial<IFormState & { resume: string }> = {};
    if (!form.name.trim())         errors.name         = 'Full name is required.';
    if (!form.email.trim())        errors.email        = 'Email is required.';
    else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(form.email))
                                   errors.email        = 'Enter a valid email address.';
    if (!form.phone.trim())        errors.phone        = 'Contact number is required.';
    if (!form.currentCtc.trim())   errors.currentCtc   = 'Current CTC is required.';
    if (!form.expectedCtc.trim())  errors.expectedCtc  = 'Expected CTC is required.';
    if (!form.noticePeriod)        errors.noticePeriod = 'Notice period is required.';
    if (!resumeFile)               errors.resume       = 'Please upload your resume.';
    setFieldErrors(errors);
    return Object.keys(errors).length === 0;
  };

  // ── Submit ──────────────────────────────────────────────────────────────────
  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    if (!validate()) {
      document.querySelector('.field-error')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      return;
    }

    setPageState('submitting');
    setSubmitError('');

    try {
      const base64 = await fileToBase64(resumeFile!);
      await submitApplication({
        jobId: Number(jobId),
        name: form.name.trim(),
        email: form.email.trim(),
        phone: form.phone.trim(),
        currentCtc: form.currentCtc.trim(),
        expectedCtc: form.expectedCtc.trim(),
        noticePeriod: form.noticePeriod,
        resumeBase64: base64,
        resumeFileName: resumeFile!.name,
        gender: form.gender || undefined,
        reasonForLeaving: form.reasonForLeaving.trim() || undefined,
        workCultureExpectation: form.workCultureExpectation.trim() || undefined,
        linkedInUrl: form.linkedInUrl.trim() || undefined,
      });
      setPageState('success');
    } catch (err) {
      setSubmitError((err as Error).message);
      setPageState('form');
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  };

  // ── Render states ───────────────────────────────────────────────────────────
  if (pageState === 'loading') {
    return (
      <div className="page-center">
        <div className="spinner" />
        <p className="loading-text">Loading job details…</p>
      </div>
    );
  }

  if (pageState === 'error') {
    return (
      <div className="error-page">
        <div className="error-card">
          <div className="error-icon">⚠️</div>
          <h2>Unable to Load Job</h2>
          <p>{loadError}</p>
        </div>
      </div>
    );
  }

  if (pageState === 'success') {
    return (
      <div className="page-center">
        <div className="success-card">
          <div className="success-icon">✓</div>
          <h2>Application Submitted!</h2>
          <p>Thank you, <strong>{form.name}</strong>. We have received your application for</p>
          <p className="success-job">{job?.title}</p>
          <p className="success-sub">
            Our HR team will review your profile and reach out to you at <strong>{form.email}</strong>.
          </p>
        </div>
      </div>
    );
  }

  // ── Main form ───────────────────────────────────────────────────────────────
  const isSubmitting = pageState === 'submitting';

  return (
    <div className="apply-page">
      {/* Header */}
      <header className="apply-header">
        <div className="apply-header-inner">
          <span className="company-name">OpATS</span>
          <span className="header-sep">|</span>
          <span className="header-label">Recruitment Portal</span>
        </div>
      </header>

      <main className="apply-main">
        {/* Job summary card */}
        <div className="job-card">
          <h1 className="job-title">{job?.title}</h1>
          <div className="job-meta">
            {job?.department  && <span className="meta-chip">🏢 {job.department}</span>}
            {job?.jobLocation && <span className="meta-chip">📍 {job.jobLocation}</span>}
            {job?.jobType     && <span className="meta-chip">🕐 {job.jobType}</span>}
            {job?.experience  && <span className="meta-chip">⏳ {job.experience}</span>}
          </div>
          {job?.requiredSkills && (
            <div className="skills-row">
              <span className="skills-label">Must Have:</span>
              {job.requiredSkills.split(',').map(s => s.trim()).filter(Boolean).map(s => (
                <span key={s} className="skill-tag must-have">{s}</span>
              ))}
            </div>
          )}
          {job?.goodToHaveSkills && (
            <div className="skills-row">
              <span className="skills-label">Good To Have:</span>
              {job.goodToHaveSkills.split(',').map(s => s.trim()).filter(Boolean).map(s => (
                <span key={s} className="skill-tag good-to-have">{s}</span>
              ))}
            </div>
          )}
          {job?.jobDescription && (
            <details className="jd-details">
              <summary>View full job description</summary>
              <pre className="jd-text">{job.jobDescription}</pre>
            </details>
          )}
        </div>

        {/* Application form */}
        <form className="apply-form" onSubmit={e => { void handleSubmit(e); }} noValidate>
          <h2 className="form-heading">Your Application</h2>

          {submitError && (
            <div className="alert alert-error">{submitError}</div>
          )}

          {/* ── Required section ─────────────────────────────────────────── */}
          <div className="section-label">Required Information</div>

          <div className="form-row">
            <div className="form-group">
              <label className="form-label required">Full Name</label>
              <input
                className={`form-input ${fieldErrors.name ? 'input-error' : ''}`}
                type="text"
                placeholder="John Smith"
                value={form.name}
                onChange={set('name')}
                disabled={isSubmitting}
                autoComplete="name"
              />
              {fieldErrors.name && <span className="field-error">{fieldErrors.name}</span>}
            </div>
            <div className="form-group">
              <label className="form-label required">Email Address</label>
              <input
                className={`form-input ${fieldErrors.email ? 'input-error' : ''}`}
                type="email"
                placeholder="john@example.com"
                value={form.email}
                onChange={set('email')}
                disabled={isSubmitting}
                autoComplete="email"
              />
              {fieldErrors.email && <span className="field-error">{fieldErrors.email}</span>}
            </div>
          </div>

          <div className="form-row">
            <div className="form-group">
              <label className="form-label required">Contact Number</label>
              <input
                className={`form-input ${fieldErrors.phone ? 'input-error' : ''}`}
                type="tel"
                placeholder="+1 555 000 0000"
                value={form.phone}
                onChange={set('phone')}
                disabled={isSubmitting}
                autoComplete="tel"
              />
              {fieldErrors.phone && <span className="field-error">{fieldErrors.phone}</span>}
            </div>
            <div className="form-group">
              <label className="form-label required">Notice Period</label>
              <select
                className={`form-input ${fieldErrors.noticePeriod ? 'input-error' : ''}`}
                value={form.noticePeriod}
                onChange={set('noticePeriod')}
                disabled={isSubmitting}
              >
                <option value="">Select…</option>
                {NOTICE_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
              {fieldErrors.noticePeriod && <span className="field-error">{fieldErrors.noticePeriod}</span>}
            </div>
          </div>

          <div className="form-row">
            <div className="form-group">
              <label className="form-label required">Current CTC</label>
              <input
                className={`form-input ${fieldErrors.currentCtc ? 'input-error' : ''}`}
                type="text"
                placeholder="e.g. $85,000 / year"
                value={form.currentCtc}
                onChange={set('currentCtc')}
                disabled={isSubmitting}
              />
              {fieldErrors.currentCtc && <span className="field-error">{fieldErrors.currentCtc}</span>}
            </div>
            <div className="form-group">
              <label className="form-label required">Expected CTC</label>
              <input
                className={`form-input ${fieldErrors.expectedCtc ? 'input-error' : ''}`}
                type="text"
                placeholder="e.g. $95,000 / year"
                value={form.expectedCtc}
                onChange={set('expectedCtc')}
                disabled={isSubmitting}
              />
              {fieldErrors.expectedCtc && <span className="field-error">{fieldErrors.expectedCtc}</span>}
            </div>
          </div>

          {/* Resume upload */}
          <div className="form-group">
            <label className="form-label required">Resume / CV</label>
            <div
              className={`file-drop ${isDragging ? 'dragging' : ''} ${resumeFile ? 'has-file' : ''} ${fieldErrors.resume ? 'input-error' : ''}`}
              onDragOver={e => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={onDrop}
              onClick={() => fileInputRef.current?.click()}
              role="button"
              tabIndex={0}
              onKeyDown={e => e.key === 'Enter' && fileInputRef.current?.click()}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".pdf,.docx,.doc"
                className="file-input-hidden"
                onChange={onFileInputChange}
                disabled={isSubmitting}
              />
              {resumeFile ? (
                <div className="file-selected">
                  <span className="file-icon">📄</span>
                  <span className="file-name">{resumeFile.name}</span>
                  <span className="file-size">({(resumeFile.size / 1024).toFixed(0)} KB)</span>
                  <button
                    type="button"
                    className="file-remove"
                    onClick={e => { e.stopPropagation(); setResumeFile(null); }}
                  >×</button>
                </div>
              ) : (
                <div className="file-prompt">
                  <span className="file-upload-icon">⬆</span>
                  <span className="file-prompt-text">
                    Drag &amp; drop or <span className="file-link">browse</span>
                  </span>
                  <span className="file-hint">PDF or DOCX · Max 5 MB</span>
                </div>
              )}
            </div>
            {(resumeError || fieldErrors.resume) && (
              <span className="field-error">{resumeError || fieldErrors.resume}</span>
            )}
          </div>

          {/* ── Optional section ─────────────────────────────────────────── */}
          <div className="section-label optional-label">Optional Information</div>

          {/* Gender */}
          <div className="form-group">
            <label className="form-label">Gender</label>
            <div className="radio-group">
              {GENDER_OPTIONS.map(g => (
                <label key={g} className="radio-label">
                  <input
                    type="radio"
                    name="gender"
                    value={g}
                    checked={form.gender === g}
                    onChange={set('gender')}
                    disabled={isSubmitting}
                  />
                  <span className="radio-text">{g}</span>
                </label>
              ))}
            </div>
          </div>

          <div className="form-group">
            <label className="form-label">LinkedIn Profile URL</label>
            <input
              className="form-input"
              type="url"
              placeholder="https://linkedin.com/in/yourprofile"
              value={form.linkedInUrl}
              onChange={set('linkedInUrl')}
              disabled={isSubmitting}
            />
          </div>

          <div className="form-group">
            <label className="form-label">Reason for Leaving Current Role</label>
            <textarea
              className="form-input form-textarea"
              placeholder="Brief explanation (optional)…"
              value={form.reasonForLeaving}
              onChange={set('reasonForLeaving')}
              rows={3}
              disabled={isSubmitting}
            />
          </div>

          <div className="form-group">
            <label className="form-label">What do you expect in a work culture?</label>
            <textarea
              className="form-input form-textarea"
              placeholder="Describe your ideal work environment…"
              value={form.workCultureExpectation}
              onChange={set('workCultureExpectation')}
              rows={3}
              disabled={isSubmitting}
            />
          </div>

          {/* Submit */}
          <div className="form-footer">
            <button
              type="submit"
              className="submit-btn"
              disabled={isSubmitting}
            >
              {isSubmitting
                ? <><span className="btn-spinner" />Submitting…</>
                : 'Submit Application'}
            </button>
            <p className="privacy-note">
              Your information is used only for this recruitment process and will not be shared externally.
            </p>
          </div>
        </form>
      </main>

      <footer className="apply-footer">
        <p>Powered by OpATS · {new Date().getFullYear()}</p>
      </footer>
    </div>
  );
}
