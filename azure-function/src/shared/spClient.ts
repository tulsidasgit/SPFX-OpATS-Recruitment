import { spfi, SPFI } from '@pnp/sp';
import { NodeFetchWithRetry, DefaultInit, MSAL } from '@pnp/nodejs';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';

let _sp: SPFI | null = null;

export function getSpClient(): SPFI {
  if (_sp) return _sp;

  const siteUrl    = process.env.SP_SITE_URL    ?? '';
  const tenantId   = process.env.SP_TENANT_ID   ?? '';
  const clientId   = process.env.SP_CLIENT_ID   ?? '';
  const clientSecret = process.env.SP_CLIENT_SECRET ?? '';

  if (!siteUrl || !tenantId || !clientId || !clientSecret) {
    throw new Error('Missing SharePoint configuration. Check SP_SITE_URL, SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET environment variables.');
  }

  _sp = spfi(siteUrl).using(
    DefaultInit(),
    NodeFetchWithRetry(),
    MSAL({
      authority: `https://login.microsoftonline.com/${tenantId}/`,
      clientId,
      clientSecret,
    })
  );

  return _sp;
}

// Server-relative root path for the Resumes library
export function getResumeFolderPath(department: string, jobTitle: string): string {
  const siteUrl = process.env.SP_SITE_URL ?? '';
  const sitePath = new URL(siteUrl).pathname.replace(/\/$/, '');
  const sanitize = (s: string): string => s.replace(/[^a-zA-Z0-9_\- ]/g, '').trim();
  return `${sitePath}/Resumes/${sanitize(department)}/${sanitize(jobTitle)}`;
}

export function corsHeaders(): Record<string, string> {
  return {
    'Access-Control-Allow-Origin': process.env.CORS_ORIGIN ?? '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  };
}
