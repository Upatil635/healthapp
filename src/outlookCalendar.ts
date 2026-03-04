import * as vscode from 'vscode';

// ─── TYPES ───────────────────────────────────────────────────
export interface OutlookMeeting {
  id: string;
  title: string;
  start: string;       // HH:MM format
  end: string;
  joinUrl?: string;    // Teams/Zoom link if available
  location?: string;
  organizer?: string;
}

interface TokenCache {
  accessToken: string;
  expiresAt: number;
}

// ─── CONFIG ─────────────────────────────────────────────────
const AZURE_CLIENT_ID  = '56f7ef10-d097-4356-a0f7-e4c4d3ae39c5';
const AZURE_TENANT_ID  = 'common'; // 'common' = Any Entra ID Tenant + Personal Microsoft accounts
const GRAPH_SCOPES     = ['Calendars.Read', 'User.Read', 'offline_access'];
const REDIRECT_URI     = 'https://login.microsoftonline.com/common/oauth2/nativeclient';

let syncTimer: NodeJS.Timeout | undefined;
let tokenCache: TokenCache | null = null;

// ─── AUTH: GET DEVICE CODE FLOW ──────────────────────────────
// Uses Device Code Flow — user opens URL and enters code
// Perfect for VS Code extensions (no browser redirect needed)

export async function signInToOutlook(context: vscode.ExtensionContext): Promise<boolean> {
  try {
    // Step 1: Request device code
    const deviceCodeUrl = `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/devicecode`;
    const scope = GRAPH_SCOPES.join(' ');

    const deviceResp = await fetch(deviceCodeUrl, {
  method: 'POST',
  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  body: new URLSearchParams({
    client_id: AZURE_CLIENT_ID,
    scope,
  }).toString(),
});

// ADD THESE DEBUG LINES:
console.log('Device code response status:', deviceResp.status);
const rawText = await deviceResp.text();
console.log('Device code raw response:', rawText);

let deviceData: any;
try {
  deviceData = JSON.parse(rawText);
} catch(e) {
  vscode.window.showErrorMessage(`❌ Azure returned non-JSON: ${rawText.substring(0, 200)}`);
  return false;
}

if (!deviceData.user_code) {
  const errMsg = deviceData.error_description || deviceData.error || rawText;
  vscode.window.showErrorMessage(`❌ Azure error: ${errMsg}`);
  return false;
} 


    // Step 2: Show user the code + URL
    const choice = await vscode.window.showInformationMessage(
      `📋 To sign in to Outlook, go to: ${deviceData.verification_uri}\n\nEnter code: ${deviceData.user_code}`,
      { modal: true },
      'Open Browser & Copy Code',
      'Cancel'
    );

    if (choice !== 'Open Browser & Copy Code') { return false; }

    // Copy code to clipboard
    await vscode.env.clipboard.writeText(deviceData.user_code);
    await vscode.env.openExternal(vscode.Uri.parse(deviceData.verification_uri));

    vscode.window.showInformationMessage(`✅ Code copied! Paste it in the browser: ${deviceData.user_code}`);

    // Step 3: Poll for token
    const token = await pollForToken(context, deviceData.device_code, deviceData.interval || 5);
    return token !== null;

  } catch (err) {
    vscode.window.showErrorMessage(`❌ Outlook sign-in error: ${err}`);
    return false;
  }
}

// ─── POLL FOR TOKEN ──────────────────────────────────────────

async function pollForToken(context: vscode.ExtensionContext, deviceCode: string, intervalSec: number): Promise<string | null> {
  const tokenUrl = `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`;
  const maxAttempts = 60; // 5 min timeout
  let attempts = 0;

  return new Promise(resolve => {
    const poll = setInterval(async () => {
      attempts++;
      if (attempts > maxAttempts) {
        clearInterval(poll);
        vscode.window.showErrorMessage('❌ Sign-in timed out. Please try again.');
        resolve(null);
        return;
      }

      try {
        const resp = await fetch(tokenUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({
            grant_type:  'urn:ietf:params:oauth:grant-type:device_code',
            client_id:   AZURE_CLIENT_ID,
            device_code: deviceCode,
          }).toString(),
        });

        const data = await resp.json() as any;

        if (data.access_token) {
          clearInterval(poll);
          // Save tokens to globalState
          await context.globalState.update('outlookAccessToken',  data.access_token);
          await context.globalState.update('outlookRefreshToken', data.refresh_token || '');
          await context.globalState.update('outlookTokenExpiry',  Date.now() + (data.expires_in * 1000));

          tokenCache = { accessToken: data.access_token, expiresAt: Date.now() + (data.expires_in * 1000) };
          vscode.window.showInformationMessage('✅ Signed in to Outlook Calendar!');
          resolve(data.access_token);
        }
        // 'authorization_pending' = still waiting, keep polling
      } catch (e) {
        // keep polling
      }
    }, intervalSec * 1000);
  });
}

// ─── GET VALID ACCESS TOKEN ──────────────────────────────────

async function getAccessToken(context: vscode.ExtensionContext): Promise<string | null> {
  // Use cached token if still valid (with 5 min buffer)
  if (tokenCache && tokenCache.expiresAt > Date.now() + 5 * 60000) {
    return tokenCache.accessToken;
  }

  // Try to refresh using saved refresh token
  const refreshToken = context.globalState.get<string>('outlookRefreshToken', '');
  if (!refreshToken) { return null; }

  try {
    const resp = await fetch(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type:    'refresh_token',
        client_id:     AZURE_CLIENT_ID,
        refresh_token: refreshToken,
        scope:         GRAPH_SCOPES.join(' '),
      }).toString(),
    });

    const data = await resp.json() as any;
    if (data.access_token) {
      await context.globalState.update('outlookAccessToken',  data.access_token);
      await context.globalState.update('outlookRefreshToken', data.refresh_token || refreshToken);
      await context.globalState.update('outlookTokenExpiry',  Date.now() + (data.expires_in * 1000));
      tokenCache = { accessToken: data.access_token, expiresAt: Date.now() + (data.expires_in * 1000) };
      return data.access_token;
    }
  } catch (e) {}

  return null;
}

// ─── SIGN OUT ────────────────────────────────────────────────

export async function signOutFromOutlook(context: vscode.ExtensionContext): Promise<void> {
  await context.globalState.update('outlookAccessToken',  undefined);
  await context.globalState.update('outlookRefreshToken', undefined);
  await context.globalState.update('outlookTokenExpiry',  undefined);
  await context.globalState.update('outlookMeetings',     undefined);
  tokenCache = null;
  stopOutlookSync();
  vscode.window.showInformationMessage('✅ Signed out from Outlook Calendar.');
}

// ─── CHECK IF SIGNED IN ──────────────────────────────────────

export function isSignedInToOutlook(context: vscode.ExtensionContext): boolean {
  const token = context.globalState.get<string>('outlookAccessToken', '');
  const expiry = context.globalState.get<number>('outlookTokenExpiry', 0);
  return !!(token && expiry > Date.now());
}

// ─── FETCH TODAY'S MEETINGS FROM GRAPH API ───────────────────

export async function fetchTodayOutlookMeetings(context: vscode.ExtensionContext): Promise<OutlookMeeting[]> {
  const token = await getAccessToken(context);
  if (!token) { return []; }

  try {
    // Build today's date range in ISO format
    const now = new Date();
    const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
    const endOfDay   = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

    const startIso = startOfDay.toISOString();
    const endIso   = endOfDay.toISOString();

    const url = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${startIso}&endDateTime=${endIso}&$select=subject,start,end,location,organizer,onlineMeeting&$orderby=start/dateTime&$top=20`;

    const resp = await fetch(url, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type':  'application/json',
      },
    });

    if (!resp.ok) {
      if (resp.status === 401) {
        // Token expired — clear cache to force refresh
        tokenCache = null;
        await context.globalState.update('outlookAccessToken', undefined);
      }
      return [];
    }

    const data = await resp.json() as any;
    const events = data.value || [];

    const meetings: OutlookMeeting[] = events.map((event: any) => {
      const startDate = new Date(event.start?.dateTime + 'Z');
      const endDate   = new Date(event.end?.dateTime   + 'Z');
      const hh = String(startDate.getHours()).padStart(2, '0');
      const mm = String(startDate.getMinutes()).padStart(2, '0');
      const eh = String(endDate.getHours()).padStart(2, '0');
      const em = String(endDate.getMinutes()).padStart(2, '0');

      return {
        id:        event.id || '',
        title:     event.subject || 'No Title',
        start:     `${hh}:${mm}`,
        end:       `${eh}:${em}`,
        joinUrl:   event.onlineMeeting?.joinUrl || '',
        location:  event.location?.displayName  || '',
        organizer: event.organizer?.emailAddress?.name || '',
      };
    });

    // Save to globalState for offline access
    await context.globalState.update('outlookMeetings', meetings);
    await context.globalState.update('outlookLastSync', new Date().toLocaleTimeString());

    return meetings;

  } catch (err) {
    console.error('Outlook fetch error:', err);
    // Return cached meetings if available
    return context.globalState.get<OutlookMeeting[]>('outlookMeetings', []);
  }
}

// ─── START AUTO-SYNC EVERY 30 MINS ───────────────────────────

export function startOutlookSync(
  context: vscode.ExtensionContext,
  onSync: (meetings: OutlookMeeting[]) => void
): void {
  stopOutlookSync();

  // Fetch immediately on start
  fetchTodayOutlookMeetings(context).then(onSync);

  // Then every 30 minutes
  syncTimer = setInterval(async () => {
    const meetings = await fetchTodayOutlookMeetings(context);
    onSync(meetings);
    vscode.window.setStatusBarMessage('📅 Outlook synced at ' + new Date().toLocaleTimeString(), 3000);
  }, 30 * 60 * 1000);
}

// ─── STOP AUTO-SYNC ──────────────────────────────────────────

export function stopOutlookSync(): void {
  if (syncTimer) {
    clearInterval(syncTimer);
    syncTimer = undefined;
  }
}

// ─── GET CACHED MEETINGS ─────────────────────────────────────

export function getCachedOutlookMeetings(context: vscode.ExtensionContext): OutlookMeeting[] {
  return context.globalState.get<OutlookMeeting[]>('outlookMeetings', []);
}

export function getLastSyncTime(context: vscode.ExtensionContext): string {
  return context.globalState.get<string>('outlookLastSync', 'Never');
}