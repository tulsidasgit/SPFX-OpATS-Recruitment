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
  Panel,
  PanelType,
  PrimaryButton,
} from '@fluentui/react';
import { CATEGORY_ORDER, CATEGORY_CONFIG, deriveCategory } from '../shared/candidateCategory';
import { SPFI } from '@pnp/sp';
import { SpService } from '../shared/SpService';
import { IJobOpening, ICandidate, IInterview } from '../shared/models';
import styles from './CompletedJobs.module.scss';

interface ICompletedJobsProps {
  sp: SPFI;
}

interface ICompletedJobsState {
  jobs: IJobOpening[];
  loadingJobs: boolean;
  jobsError: string;
  expandedJobIds: Set<number>;
  candidatesMap: Record<number, ICandidate[]>;
  interviewsMap: Record<number, IInterview[]>;
  loadingDataFor: Set<number>;
  reportJobId: number | undefined;
  loadingReport: boolean;
}

function getScoreColor(score: number): string {
  if (score >= 75) return '#107c10';
  if (score >= 50) return '#f7630c';
  return '#a80000';
}

function getRecommendationClass(recommendation: string): string {
  switch (recommendation) {
    case 'Recommended':     return styles.badgeGreen;
    case 'Maybe':           return styles.badgeOrange;
    case 'Not Recommended': return styles.badgeRed;
    default:                return styles.badgeGrey;
  }
}

function getOutcome(candidate: ICandidate): { label: string; cssClass: string } {
  if (candidate.applicationStatus === 'Rejected') {
    return { label: 'Rejected', cssClass: styles.badgeRed };
  }
  if (candidate.recommendation === 'Recommended') {
    return { label: 'Selected', cssClass: styles.badgeGreen };
  }
  return { label: 'Not Selected', cssClass: styles.badgeGrey };
}

function renderSkillChips(skills: string, chipClass: string): React.ReactNode {
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

function csvEscape(value: string): string {
  if (/[",\n\r]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

function summarizeInterviews(candidateId: number, interviews: IInterview[]): string {
  const mine = interviews
    .filter(iv => iv.candidateId === candidateId)
    .sort((a, b) => parseInt(a.interviewRound, 10) - parseInt(b.interviewRound, 10));

  if (mine.length === 0) return 'No interviews';

  return mine
    .map(iv => {
      const date = iv.scheduledDate ? new Date(iv.scheduledDate).toLocaleDateString('en-GB') : '—';
      return `Round ${iv.interviewRound}: ${iv.interviewerEmail} on ${date} — ${iv.feedbackStatus}`;
    })
    .join('; ');
}

function buildCsv(job: IJobOpening, candidates: ICandidate[], interviews: IInterview[]): string {
  const header = [
    'Candidate Name', 'Email', 'Phone', 'Outcome', 'Application Status',
    'Fitment Score', 'Recommendation', 'Experience Match',
    'Matching Skills', 'Missing Skills', 'AI Summary', 'HR Feedback', 'Interviews',
  ];

  const rows = candidates.map(c => [
    c.candidateName,
    c.email,
    c.phone,
    getOutcome(c).label,
    c.applicationStatus ?? '',
    String(c.fitmentScore),
    c.recommendation,
    c.experienceMatch,
    c.matchingSkills,
    c.missingSkills,
    c.aiSummary,
    c.hrFeedback,
    summarizeInterviews(c.id, interviews),
  ]);

  return [header, ...rows]
    .map(row => row.map(csvEscape).join(','))
    .join('\r\n');
}

function downloadCsv(filename: string, csv: string): void {
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function sanitizeFilenamePart(value: string): string {
  return value.replace(/[\\/:*?"<>|]/g, '-').trim();
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function fmtDate(iso: string | undefined): string {
  if (!iso) return '—';
  return new Date(iso).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

function scoreHex(score: number): string {
  if (score >= 75) return '#107c10';
  if (score >= 50) return '#f7630c';
  return score > 0 ? '#a80000' : '#605e5c';
}

function recColor(rec: string): { bg: string; fg: string } {
  if (rec === 'Recommended')     return { bg: '#dff6dd', fg: '#107c10' };
  if (rec === 'Maybe')           return { bg: '#fff4ce', fg: '#8a8000' };
  if (rec === 'Not Recommended') return { bg: '#fde7e9', fg: '#a80000' };
  return { bg: '#edebe9', fg: '#605e5c' };
}

function buildReportHtml(job: IJobOpening, candidates: ICandidate[], interviews: IInterview[]): string {
  const generatedAt = new Date().toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });

  // Pipeline counts
  const stageCounts: Partial<Record<string, number>> = {};
  candidates.forEach(c => {
    const cat = deriveCategory(c, interviews);
    stageCounts[cat] = (stageCounts[cat] ?? 0) + 1;
  });

  const pipelineCardsHtml = CATEGORY_ORDER.map(cat => {
    const count = stageCounts[cat] ?? 0;
    const { label, color } = CATEGORY_CONFIG[cat];
    return `
      <div style="background:${color};color:#fff;border-radius:8px;padding:14px 18px;min-width:90px;text-align:center;flex:1">
        <div style="font-size:30px;font-weight:800;line-height:1">${count}</div>
        <div style="font-size:10px;text-transform:uppercase;letter-spacing:0.5px;margin-top:5px;opacity:0.9;font-weight:600">${escHtml(label)}</div>
      </div>`;
  }).join('');

  // Skills chips inline helper
  const skillChips = (skills: string, bg: string, fg: string): string =>
    skills
      ? skills.split(',').map(s => s.trim()).filter(Boolean)
          .map(s => `<span style="display:inline-block;background:${bg};color:${fg};padding:1px 7px;border-radius:8px;font-size:10px;font-weight:600;margin:1px 2px 1px 0">${escHtml(s)}</span>`)
          .join('')
      : '<span style="color:#605e5c;font-size:11px">—</span>';

  // Candidate rows
  const candidateRowsHtml = candidates.map((c, idx) => {
    const stage = deriveCategory(c, interviews);
    const { color: stageColor } = CATEGORY_CONFIG[stage];
    const { label: stageLabel } = CATEGORY_CONFIG[stage];
    const sc = scoreHex(c.fitmentScore);
    const rc = recColor(c.recommendation);
    const mine = interviews
      .filter(iv => iv.candidateId === c.id)
      .sort((a, b) => parseInt(a.interviewRound, 10) - parseInt(b.interviewRound, 10));

    const ivHtml = mine.length === 0
      ? '<span style="color:#605e5c;font-size:11px">None</span>'
      : mine.map(iv => {
          const ivDate = iv.scheduledDate ? new Date(iv.scheduledDate).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';
          const fbBg = iv.feedbackStatus === 'Submitted' ? '#dff6dd' : '#fff4ce';
          const fbFg = iv.feedbackStatus === 'Submitted' ? '#107c10' : '#8a8000';
          return `<div style="margin-bottom:3px;font-size:11px">
            <span style="background:#fd800b;color:#fff;padding:1px 6px;border-radius:6px;font-weight:700;font-size:10px">Rd ${iv.interviewRound}</span>
            &nbsp;${escHtml(iv.interviewerEmail)}&nbsp;·&nbsp;${ivDate}
            <span style="background:${fbBg};color:${fbFg};padding:1px 6px;border-radius:6px;font-weight:600;font-size:10px;margin-left:4px">${iv.feedbackStatus}</span>
          </div>`;
        }).join('');

    const rowBg = idx % 2 === 0 ? '#ffffff' : '#f9f8f7';

    return `
      <tr style="background:${rowBg}">
        <td style="padding:10px 12px;border-bottom:1px solid #edebe9;vertical-align:top">
          <div style="font-weight:600;font-size:13px;color:#323130">${escHtml(c.candidateName)}</div>
          <div style="font-size:11px;color:#605e5c;margin-top:2px">${escHtml(c.email)}${c.phone ? ' · ' + escHtml(c.phone) : ''}</div>
          ${c.hrFeedback ? `<div style="margin-top:5px;padding:4px 7px;background:#fff4ce;border-left:3px solid #f7630c;border-radius:0 3px 3px 0;font-size:11px;color:#323130"><strong style="color:#ca5010">HR:</strong> ${escHtml(c.hrFeedback)}</div>` : ''}
        </td>
        <td style="padding:10px 12px;border-bottom:1px solid #edebe9;vertical-align:top">
          <span style="background:${stageColor};color:#fff;padding:2px 9px;border-radius:10px;font-size:11px;font-weight:700;white-space:nowrap">${escHtml(stageLabel)}</span>
        </td>
        <td style="padding:10px 12px;border-bottom:1px solid #edebe9;vertical-align:top">
          <span style="background:${sc};color:#fff;padding:3px 10px;border-radius:10px;font-size:12px;font-weight:800;display:inline-block">${c.fitmentScore > 0 ? c.fitmentScore + '%' : '—'}</span>
          <div style="margin-top:6px">${skillChips(c.matchingSkills, '#dff6dd', '#107c10')}</div>
          <div style="margin-top:3px">${skillChips(c.missingSkills, '#fde7e9', '#a80000')}</div>
        </td>
        <td style="padding:10px 12px;border-bottom:1px solid #edebe9;vertical-align:top">
          ${c.recommendation ? `<span style="background:${rc.bg};color:${rc.fg};padding:2px 9px;border-radius:10px;font-size:11px;font-weight:600">${escHtml(c.recommendation)}</span>` : '<span style="color:#605e5c;font-size:11px">—</span>'}
          ${c.aiSummary ? `<div style="margin-top:6px;font-size:11px;color:#323130;line-height:1.5">${escHtml(c.aiSummary)}</div>` : ''}
        </td>
        <td style="padding:10px 12px;border-bottom:1px solid #edebe9;vertical-align:top">${ivHtml}</td>
      </tr>`;
  }).join('');

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>Recruitment Report — ${escHtml(job.title)}</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',Arial,sans-serif;color:#323130;background:#fff;font-size:13px}
    table{border-collapse:collapse;width:100%}
    @media print{
      body{-webkit-print-color-adjust:exact;print-color-adjust:exact}
      .no-print{display:none!important}
      .page-break{page-break-before:always}
    }
  </style>
</head>
<body>
  <!-- Print toolbar -->
  <div class="no-print" style="background:#f3f2f1;padding:10px 40px;display:flex;align-items:center;gap:12px;border-bottom:1px solid #edebe9">
    <button onclick="window.print()" style="background:#0078d4;color:#fff;border:none;padding:8px 22px;border-radius:4px;font-size:14px;font-weight:600;cursor:pointer">
      ⬇ Print / Save as PDF
    </button>
    <span style="font-size:12px;color:#605e5c">Use your browser&apos;s Print dialog → &quot;Save as PDF&quot; to download</span>
  </div>

  <!-- Report header -->
  <div style="background:#0078d4;padding:28px 40px;position:relative">
    <div style="font-size:11px;color:#c7e0f4;font-weight:600;letter-spacing:1px;text-transform:uppercase">OpATS — Recruitment Report</div>
    <div style="font-size:22px;font-weight:700;color:#fff;margin-top:6px">${escHtml(job.title)}</div>
    <div style="font-size:13px;color:#c7e0f4;margin-top:4px">${escHtml(job.department)}${job.jobLocation ? ' · ' + escHtml(job.jobLocation) : ''}</div>
    <div style="position:absolute;top:28px;right:40px;font-size:11px;color:#c7e0f4;font-weight:600;letter-spacing:1px">CONFIDENTIAL &nbsp;|&nbsp; Generated ${generatedAt}</div>
  </div>

  <!-- Job Overview -->
  <div style="padding:24px 40px;border-bottom:1px solid #edebe9">
    <div style="font-size:12px;font-weight:700;color:#0078d4;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:16px">Job Overview</div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px">
      ${[
        ['Job Title',    job.jobTitle || job.title],
        ['Department',   job.department],
        ['Location',     job.jobLocation || '—'],
        ['Job Type',     job.jobType || '—'],
        ['Experience',   job.experience || '—'],
        ['Posted By',    job.postedBy || '—'],
        ['Created',      fmtDate(job.created)],
        ['Deadline',     fmtDate(job.dueDate)],
        ['Completed',    fmtDate(job.modified)],
      ].map(([lbl, val]) => `
        <div>
          <div style="font-size:10px;color:#605e5c;text-transform:uppercase;letter-spacing:0.5px">${lbl}</div>
          <div style="font-size:13px;color:#323130;font-weight:500;margin-top:3px">${escHtml(String(val))}</div>
        </div>`).join('')}
    </div>
  </div>

  <!-- Pipeline Summary -->
  <div style="padding:24px 40px;border-bottom:1px solid #edebe9">
    <div style="font-size:12px;font-weight:700;color:#0078d4;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:16px">
      Pipeline Summary &nbsp;<span style="font-weight:400;color:#605e5c;text-transform:none;letter-spacing:0">${candidates.length} candidate${candidates.length !== 1 ? 's' : ''} total</span>
    </div>
    <div style="display:flex;gap:10px;flex-wrap:wrap">${pipelineCardsHtml}</div>
  </div>

  <!-- Candidate Details -->
  <div style="padding:24px 40px 32px">
    <div style="font-size:12px;font-weight:700;color:#0078d4;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:16px">Candidate Details</div>
    <table>
      <thead>
        <tr style="background:#f3f2f1">
          <th style="padding:9px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;color:#605e5c;font-weight:700">Candidate</th>
          <th style="padding:9px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;color:#605e5c;font-weight:700">Stage</th>
          <th style="padding:9px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;color:#605e5c;font-weight:700">Score &amp; Skills</th>
          <th style="padding:9px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;color:#605e5c;font-weight:700">AI Assessment</th>
          <th style="padding:9px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;color:#605e5c;font-weight:700">Interviews</th>
        </tr>
      </thead>
      <tbody>
        ${candidateRowsHtml || `<tr><td colspan="5" style="padding:16px 12px;color:#605e5c;text-align:center">No candidates recorded for this job.</td></tr>`}
      </tbody>
    </table>
  </div>

  <!-- Footer -->
  <div style="background:#f3f2f1;padding:14px 40px;display:flex;justify-content:space-between;font-size:11px;color:#605e5c;border-top:1px solid #edebe9">
    <span>OpATS &mdash; Confidential Recruitment Report &mdash; ${escHtml(job.title)}</span>
    <span>Generated: ${generatedAt}</span>
  </div>
</body>
</html>`;
}

export class CompletedJobs extends React.Component<ICompletedJobsProps, ICompletedJobsState> {
  private _spService: SpService;

  constructor(props: ICompletedJobsProps) {
    super(props);
    this._spService = new SpService(props.sp);
    this.state = {
      jobs: [],
      loadingJobs: true,
      jobsError: '',
      expandedJobIds: new Set(),
      candidatesMap: {},
      interviewsMap: {},
      loadingDataFor: new Set(),
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
      const jobs = await this._spService.getClosedJobOpenings();
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

    // Only fetch candidates/interviews once per job
    if (this.state.candidatesMap[jobId] !== undefined) return;

    const loading = new Set(this.state.loadingDataFor);
    loading.add(jobId);
    this.setState({ loadingDataFor: loading });

    try {
      const [candidates, interviews] = await Promise.all([
        this._spService.getCandidatesByJobOpening(jobId),
        this._spService.getInterviewsByJobOpening(jobId),
      ]);
      this.setState(prev => ({
        candidatesMap: { ...prev.candidatesMap, [jobId]: candidates },
        interviewsMap: { ...prev.interviewsMap, [jobId]: interviews },
        loadingDataFor: new Set(Array.from(prev.loadingDataFor).filter(id => id !== jobId)),
      }));
    } catch {
      this.setState(prev => ({
        loadingDataFor: new Set(Array.from(prev.loadingDataFor).filter(id => id !== jobId)),
      }));
    }
  }

  private _onExportCsv = async (job: IJobOpening): Promise<void> => {
    const cachedCandidates = this.state.candidatesMap[job.id];
    const cachedInterviews = this.state.interviewsMap[job.id];

    let candidates: ICandidate[];
    let interviews: IInterview[];

    if (cachedCandidates === undefined || cachedInterviews === undefined) {
      const fetched = await Promise.all([
        this._spService.getCandidatesByJobOpening(job.id),
        this._spService.getInterviewsByJobOpening(job.id),
      ]);
      candidates = fetched[0];
      interviews = fetched[1];
      this.setState(prev => ({
        candidatesMap: { ...prev.candidatesMap, [job.id]: candidates },
        interviewsMap: { ...prev.interviewsMap, [job.id]: interviews },
      }));
    } else {
      candidates = cachedCandidates;
      interviews = cachedInterviews;
    }

    const csv = buildCsv(job, candidates, interviews);
    const filename = `${sanitizeFilenamePart(job.title)}-candidates.csv`;
    downloadCsv(filename, csv);
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

  private _downloadPdf = (job: IJobOpening, candidates: ICandidate[], interviews: IInterview[]): void => {
    const html = buildReportHtml(job, candidates, interviews);
    const win = window.open('', '_blank', 'width=1060,height=800,toolbar=0,menubar=0,scrollbars=1');
    if (!win) return;
    win.document.open();
    win.document.write(html);
    win.document.close();
    setTimeout(() => { win.print(); }, 600);
  };

  private _renderReportPanel(): React.ReactNode {
    const { reportJobId, jobs, candidatesMap, interviewsMap, loadingReport } = this.state;
    if (reportJobId === undefined) return null;

    const job = jobs.find(j => j.id === reportJobId);
    if (!job) return null;

    const candidates = candidatesMap[reportJobId] ?? [];
    const interviews = interviewsMap[reportJobId] ?? [];

    const stageCounts: Partial<Record<string, number>> = {};
    candidates.forEach(c => {
      const cat = deriveCategory(c, interviews);
      stageCounts[cat] = (stageCounts[cat] ?? 0) + 1;
    });

    return (
      <Panel
        isOpen
        type={PanelType.medium}
        headerText={`Report — ${job.title}`}
        onDismiss={() => this.setState({ reportJobId: undefined })}
      >
        <div style={{ padding: '16px 0 32px' }}>
          {loadingReport && <Spinner size={SpinnerSize.medium} label="Loading report data…" />}

          {!loadingReport && (
            <>
              {/* Download PDF button */}
              <div style={{ marginBottom: 20 }}>
                <PrimaryButton
                  text="Download PDF"
                  iconProps={{ iconName: 'PDF' }}
                  onClick={() => this._downloadPdf(job, candidates, interviews)}
                />
              </div>

              {/* Job Overview */}
              <div style={{ background: '#f3f2f1', borderRadius: 6, padding: '14px 16px', marginBottom: 20 }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: '#0078d4', textTransform: 'uppercase', letterSpacing: '0.6px', marginBottom: 12 }}>Job Overview</div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px 16px' }}>
                  {([
                    ['Job Title',   job.jobTitle || job.title],
                    ['Department',  job.department],
                    ['Location',    job.jobLocation || '—'],
                    ['Job Type',    job.jobType || '—'],
                    ['Experience',  job.experience || '—'],
                    ['Posted By',   job.postedBy || '—'],
                    ['Created',     fmtDate(job.created)],
                    ['Deadline',    fmtDate(job.dueDate)],
                    ['Completed',   fmtDate(job.modified)],
                  ] as [string, string][]).map(([lbl, val]) => (
                    <div key={lbl}>
                      <div style={{ fontSize: 10, color: '#605e5c', textTransform: 'uppercase', letterSpacing: '0.5px' }}>{lbl}</div>
                      <div style={{ fontSize: 13, color: '#323130', fontWeight: 500, marginTop: 2 }}>{val}</div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Pipeline Summary */}
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: '#0078d4', textTransform: 'uppercase', letterSpacing: '0.6px', marginBottom: 10 }}>
                  Pipeline Summary &nbsp;
                  <span style={{ fontWeight: 400, color: '#605e5c', textTransform: 'none', letterSpacing: 0 }}>
                    {candidates.length} candidate{candidates.length !== 1 ? 's' : ''} total
                  </span>
                </div>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                  {CATEGORY_ORDER.map(cat => {
                    const count = stageCounts[cat] ?? 0;
                    const { label, color } = CATEGORY_CONFIG[cat];
                    return (
                      <div key={cat} style={{ background: color, color: '#fff', borderRadius: 6, padding: '10px 14px', minWidth: 70, textAlign: 'center' }}>
                        <div style={{ fontSize: 22, fontWeight: 800, lineHeight: 1 }}>{count}</div>
                        <div style={{ fontSize: 10, textTransform: 'uppercase', marginTop: 4, opacity: 0.9, fontWeight: 600 }}>{label}</div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Candidate list */}
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, color: '#0078d4', textTransform: 'uppercase', letterSpacing: '0.6px', marginBottom: 10 }}>Candidates</div>
                {candidates.length === 0 ? (
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>No candidates recorded.</Text>
                ) : (
                  candidates.map(c => {
                    const stage = deriveCategory(c, interviews);
                    const { label: stageLabel, color: stageColor } = CATEGORY_CONFIG[stage];
                    const sc = scoreHex(c.fitmentScore);
                    const rc = recColor(c.recommendation);
                    return (
                      <div key={c.id} style={{ padding: '10px 12px', borderRadius: 4, marginBottom: 6, background: '#faf9f8', border: '1px solid #edebe9' }}>
                        <div style={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
                          <div style={{ flex: 1, minWidth: 100 }}>
                            <div style={{ fontWeight: 600, fontSize: 13 }}>{c.candidateName}</div>
                            <div style={{ fontSize: 11, color: '#605e5c' }}>{c.email}</div>
                          </div>
                          <span style={{ background: stageColor, color: '#fff', padding: '2px 9px', borderRadius: 10, fontSize: 11, fontWeight: 700 }}>{stageLabel}</span>
                          <span style={{ background: sc, color: '#fff', padding: '2px 9px', borderRadius: 10, fontSize: 12, fontWeight: 800 }}>
                            {c.fitmentScore > 0 ? `${c.fitmentScore}%` : '—'}
                          </span>
                          {c.recommendation && (
                            <span style={{ background: rc.bg, color: rc.fg, padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600 }}>
                              {c.recommendation}
                            </span>
                          )}
                        </div>
                        {c.hrFeedback && (
                          <div style={{ marginTop: 6, padding: '5px 8px', background: '#fff4ce', borderLeft: '3px solid #f7630c', borderRadius: '0 3px 3px 0', fontSize: 12 }}>
                            <span style={{ fontWeight: 600, color: '#ca5010', fontSize: 11 }}>HR: </span>{c.hrFeedback}
                          </div>
                        )}
                      </div>
                    );
                  })
                )}
              </div>
            </>
          )}
        </div>
      </Panel>
    );
  }

  private _renderCandidateRow(candidate: ICandidate, interviews: IInterview[]): React.ReactNode {
    const outcome = getOutcome(candidate);
    const mine = interviews
      .filter(iv => iv.candidateId === candidate.id)
      .sort((a, b) => parseInt(a.interviewRound, 10) - parseInt(b.interviewRound, 10));

    return (
      <div key={candidate.id} className={styles.candidateRow}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 8 }}>
          <Stack tokens={{ childrenGap: 2 }} styles={{ root: { flex: 1, minWidth: 180 } }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              {candidate.candidateName}
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              {candidate.email}{candidate.phone ? ` · ${candidate.phone}` : ''}
            </Text>
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 6 }}>
            <span className={`${styles.badge} ${outcome.cssClass}`}>{outcome.label}</span>
            <span className={`${styles.badge} ${getRecommendationClass(candidate.recommendation)}`}>
              {candidate.recommendation || 'Pending'}
            </span>
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

        {candidate.matchingSkills && renderSkillChips(candidate.matchingSkills, styles.chipGreen)}
        {candidate.missingSkills && renderSkillChips(candidate.missingSkills, styles.chipRed)}

        {candidate.aiSummary && (
          <div className={styles.aiSummary}>
            <Text variant="xSmall" styles={{ root: { color: '#0078d4', fontWeight: 600 } }}>AI Summary</Text>
            <Text variant="small">{candidate.aiSummary}</Text>
          </div>
        )}

        {candidate.hrFeedback && (
          <div style={{ marginTop: 8 }}>
            <Text variant="xSmall" styles={{ root: { color: '#605e5c', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' } }}>
              HR Feedback
            </Text>
            <Text variant="small">{candidate.hrFeedback}</Text>
          </div>
        )}

        {mine.length > 0 && (
          <div className={styles.interviewHistory}>
            <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
              Interview History
            </Text>
            {mine.map(iv => (
              <div key={iv.id} className={styles.interviewRow}>
                <span className={styles.roundBadge}>Rd {iv.interviewRound}</span>
                <Text variant="small" styles={{ root: { flex: 1 } }}>{iv.interviewerEmail}</Text>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  {iv.scheduledDate ? new Date(iv.scheduledDate).toLocaleString('en-GB') : '—'}
                </Text>
                <span className={`${styles.chip} ${iv.feedbackStatus === 'Submitted' ? styles.chipGreen : styles.chipGrey}`}>
                  {iv.feedbackStatus}
                </span>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  private _renderJobBody(job: IJobOpening): React.ReactNode {
    const { candidatesMap, interviewsMap, loadingDataFor } = this.state;
    const candidates = candidatesMap[job.id] ?? [];
    const interviews = interviewsMap[job.id] ?? [];
    const isLoading = loadingDataFor.has(job.id);

    if (isLoading) {
      return <Spinner size={SpinnerSize.medium} label="Loading candidates…" styles={{ root: { padding: 16 } }} />;
    }

    return (
      <>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { padding: '12px 0 4px' } }}>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            {candidates.length} candidate{candidates.length !== 1 ? 's' : ''}
          </Text>
          <ActionButton
            iconProps={{ iconName: 'ExcelDocument' }}
            text="Export CSV"
            onClick={() => { this._onExportCsv(job).catch(() => undefined); }}
            disabled={candidates.length === 0}
          />
        </Stack>

        {candidates.length === 0 ? (
          <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block', padding: '4px 0 12px' } }}>
            No candidates were recorded for this job opening.
          </Text>
        ) : (
          candidates.map(c => this._renderCandidateRow(c, interviews))
        )}
      </>
    );
  }

  private _renderJobCard(job: IJobOpening): React.ReactNode {
    const isExpanded = this.state.expandedJobIds.has(job.id);

    return (
      <div key={job.id} className={styles.jobCard}>
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
            <span className={styles.statusBadge}>{job.status}</span>
            {job.dueDate && (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Closed by: {new Date(job.dueDate).toLocaleDateString('en-GB')}
              </Text>
            )}
            <ActionButton
              iconProps={{ iconName: 'ReportDocument' }}
              text="Report"
              onClick={e => {
                e.stopPropagation();
                this._openReport(job.id);
              }}
              styles={{ root: { height: 'auto' } }}
            />
            <ActionButton
              iconProps={{ iconName: isExpanded ? 'ChevronUp' : 'ChevronDown' }}
              styles={{ root: { padding: 0, height: 'auto', minWidth: 'auto' } }}
              ariaLabel={isExpanded ? 'Collapse' : 'Expand'}
            />
          </Stack>
        </div>

        {isExpanded && (
          <div className={styles.jobCardBody}>
            {this._renderJobBody(job)}
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
          <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>Completed Jobs</Text>
          <ActionButton
            iconProps={{ iconName: 'Refresh' }}
            text="Refresh"
            onClick={() => { this._loadJobs().catch(() => undefined); }}
            disabled={loadingJobs}
          />
        </Stack>

        {loadingJobs && <Spinner size={SpinnerSize.medium} label="Loading closed job openings…" />}

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
            No closed job openings yet. Jobs appear here once their status is set to &quot;Closed&quot;.
          </MessageBar>
        )}

        {jobs.map(job => this._renderJobCard(job))}

        {this._renderReportPanel()}
      </div>
    );
  }
}
