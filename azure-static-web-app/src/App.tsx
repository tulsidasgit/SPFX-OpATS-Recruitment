import React from 'react';
import Apply from './Apply';

export default function App(): React.ReactElement {
  const params = new URLSearchParams(window.location.search);
  const jobId = params.get('jobId') ?? '';

  if (!jobId) {
    return (
      <div className="error-page">
        <div className="error-card">
          <div className="error-icon">🔗</div>
          <h2>Invalid Application Link</h2>
          <p>This link is missing the job reference. Please use the link from the job posting.</p>
        </div>
      </div>
    );
  }

  return <Apply jobId={jobId} />;
}
