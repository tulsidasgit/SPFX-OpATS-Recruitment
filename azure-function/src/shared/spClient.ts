import { spfi, SPFI } from '@pnp/sp';
import { NodeFetchWithRetry } from '@pnp/nodejs';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';

// ACS token cache — tokens are valid ~60 min; we refresh 2 min early
const tokenCache: { value: string; expiry: number } = { value: '', expiry: 0 };

async function getACSToken(siteUrl: string, clientId: string, clientSecret: string): Promise<string> {
  if (tokenCache.value && Date.now() < tokenCache.expiry) return tokenCache.value;

  // SharePoint responds to an unauthenticated call with a 401 containing the realm in WWW-Authenticate
  const challengeRes = await fetch(`${siteUrl}/_vti_bin/client.svc`, {
    headers: { Authorization: 'Bearer' },
  });
  const wwwAuth = challengeRes.headers.get('www-authenticate') ?? '';
  const realmMatch = wwwAuth.match(/realm="([^"]+)"/);
  if (!realmMatch) {
    throw new Error(`Cannot read realm from SharePoint. Verify SP_SITE_URL is correct. WWW-Authenticate: ${wwwAuth}`);
  }
  const realm = realmMatch[1];
  const host = new URL(siteUrl).hostname;

  const tokenRes = await fetch(`https://accounts.accesscontrol.windows.net/${realm}/tokens/OAuth/2`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: `${clientId}@${realm}`,
      client_secret: clientSecret,
      resource: `00000003-0000-0ff1-ce00-000000000000/${host}@${realm}`,
    }).toString(),
  });

  if (!tokenRes.ok) {
    const body = await tokenRes.text();
    throw new Error(`ACS token request failed (${tokenRes.status}): ${body}`);
  }

  const json = await tokenRes.json() as { access_token: string; expires_in: string };
  tokenCache.value = json.access_token;
  tokenCache.expiry = Date.now() + (Number(json.expires_in) - 120) * 1000;
  return tokenCache.value;
}

// PnP.js v3 behavior: intercepts the auth event and injects the ACS Bearer token
function acsAuthBehavior(siteUrl: string, clientId: string, clientSecret: string) {
  return (instance: any) => {
    instance.on.auth.replace(async (url: URL, init: RequestInit): Promise<[URL, RequestInit]> => {
      const token = await getACSToken(siteUrl, clientId, clientSecret);
      init.headers = Object.assign({}, init.headers, { Authorization: `Bearer ${token}` });
      return [url, init];
    });
    return instance;
  };
}

let _sp: SPFI | null = null;

export function getSpClient(): SPFI {
  if (_sp) return _sp;

  const siteUrl      = process.env.SP_SITE_URL      ?? '';
  const clientId     = process.env.SP_CLIENT_ID     ?? '';
  const clientSecret = process.env.SP_CLIENT_SECRET ?? '';

  if (!siteUrl || !clientId || !clientSecret) {
    throw new Error(
      'Missing SharePoint config. Set SP_SITE_URL, SP_CLIENT_ID, SP_CLIENT_SECRET in local.settings.json.'
    );
  }

  _sp = spfi(siteUrl).using(
    NodeFetchWithRetry(),
    acsAuthBehavior(siteUrl, clientId, clientSecret)
  );

  return _sp;
}

export function getResumeFolderPath(department: string, jobTitle: string): string {
  const sitePath = new URL(process.env.SP_SITE_URL ?? '').pathname.replace(/\/$/, '');
  const clean = (s: string) => s.replace(/[^a-zA-Z0-9_\- ]/g, '').trim();
  return `${sitePath}/Resumes/${clean(department)}/${clean(jobTitle)}`;
}

export function corsHeaders(): Record<string, string> {
  return {
    'Access-Control-Allow-Origin': process.env.CORS_ORIGIN ?? '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  };
}
