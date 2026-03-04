import * as vscode from 'vscode';
import {
  signInToOutlook,
  signOutFromOutlook,
  isSignedInToOutlook,
  startOutlookSync,
  stopOutlookSync,
  fetchTodayOutlookMeetings,
  getCachedOutlookMeetings,
  getLastSyncTime,
  OutlookMeeting,
} from './outlookCalendar';

// ─── GLOBAL STATE ───────────────────────────────────────────
let timers: NodeJS.Timeout[] = [];
let meetings: { title: string; time: string; timeout: NodeJS.Timeout }[] = [];
let isRunning = false;
let statusBarItem: vscode.StatusBarItem;
let sidebarProvider: HealthSidebarProvider;
let globalContext: vscode.ExtensionContext;

interface ReminderSettings {
  teaInterval: number;
  waterInterval: number;
  stretchInterval: number;
  eyeInterval: number;
  dayStartHour: number;
  dayStartMin: number;
  snoozeMinutes: number;
}

interface SnoozedReminder {
  title: string;
  message: string;
  color: string;
  snoozeUntil: number;
}

// ─── SECURITY: HTML SANITIZATION ────────────────────────────
function escapeHtml(text: string): string {
  const map: { [key: string]: string } = {
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, c => map[c]);
}

function sanitizeColor(color: string): string {
  if (!/^(#[0-9A-Fa-f]{3}([0-9A-Fa-f]{3})?|rgb\(\d+,\s*\d+,\s*\d+\))$/.test(color)) {
    return '#0078d4';
  }
  return color;
}

function getSettings(): ReminderSettings {
  return globalContext.globalState.get<ReminderSettings>('reminderSettings', {
    teaInterval: 60, waterInterval: 30, stretchInterval: 90, eyeInterval: 20,
    dayStartHour: 9, dayStartMin: 0, snoozeMinutes: 10,
  });
}

function saveSettings(s: ReminderSettings) {
  globalContext.globalState.update('reminderSettings', s);
}

// ─── SIDEBAR PROVIDER ───────────────────────────────────────

class HealthSidebarProvider implements vscode.WebviewViewProvider {
  public static readonly viewType = 'healthReminder.sidebar';
  private _view?: vscode.WebviewView;
  private _context: vscode.ExtensionContext;

  constructor(context: vscode.ExtensionContext) {
    this._context = context;
  }

  resolveWebviewView(webviewView: vscode.WebviewView) {
    this._view = webviewView;
    webviewView.webview.options = { enableScripts: true };
    webviewView.webview.html = this.getSidebarHTML();

    webviewView.webview.onDidReceiveMessage(async message => {
      switch (message.command) {
        case 'addMeeting':
          scheduleMeeting(this._context, message.title, message.time);
          this.refresh();
          break;
        case 'deleteMeeting':
          deleteMeeting(message.index, this._context);
          this.refresh();
          break;
        case 'start':
          if (!isRunning) {
            startReminders(this._context);
            isRunning = true;
            this._context.globalState.update('isRunning', true);
            statusBarItem.text = '$(heart) Health ON';
            vscode.window.showInformationMessage('✅ Health Reminders Started!');
          }
          this.refresh();
          break;
        case 'stop':
          stopReminders();
          isRunning = false;
          this._context.globalState.update('isRunning', false);
          statusBarItem.text = '$(heart) Health OFF';
          this.refresh();
          break;
        case 'saveSettings': {
          const s: ReminderSettings = {
            teaInterval:     Math.max(1, Math.min(480, parseInt(message.teaInterval,     10) || 60)),
            waterInterval:   Math.max(1, Math.min(480, parseInt(message.waterInterval,   10) || 30)),
            stretchInterval: Math.max(1, Math.min(480, parseInt(message.stretchInterval, 10) || 90)),
            eyeInterval:     Math.max(1, Math.min(480, parseInt(message.eyeInterval,     10) || 20)),
            dayStartHour:    Math.max(0, Math.min(23,  parseInt(message.dayStartHour,    10) || 9)),
            dayStartMin:     Math.max(0, Math.min(59,  parseInt(message.dayStartMin,     10) || 0)),
            snoozeMinutes:   Math.max(1, Math.min(60,  parseInt(message.snoozeMinutes,   10) || 10)),
          };
          saveSettings(s);
          if (isRunning) {
            stopReminders();
            startReminders(this._context);
            isRunning = true;
            this._context.globalState.update('isRunning', true);
          }
          vscode.window.showInformationMessage('✅ Settings saved & reminders restarted!');
          this.refresh();
          break;
        }

        // ── Outlook commands ──
        case 'outlookSignIn': {
          const ok = await signInToOutlook(this._context);
          if (ok) {
            startOutlookSync(this._context, (mtgs) => {
              scheduleOutlookMeetings(this._context, mtgs);
              sidebarProvider.refresh();
            });
          }
          this.refresh();
          break;
        }
        case 'outlookSignOut':
          await signOutFromOutlook(this._context);
          this.refresh();
          break;
        case 'outlookSync': {
          vscode.window.setStatusBarMessage('🔄 Syncing Outlook...', 2000);
          const mtgs = await fetchTodayOutlookMeetings(this._context);
          scheduleOutlookMeetings(this._context, mtgs);
          this.refresh();
          vscode.window.showInformationMessage(`✅ Synced ${mtgs.length} meeting(s) from Outlook`);
          break;
        }
        case 'openJoinUrl':
          if (message.url) {
            vscode.env.openExternal(vscode.Uri.parse(message.url));
          }
          break;
      }
    });
  }

  refresh() {
    if (this._view) {
      this._view.webview.html = this.getSidebarHTML();
    }
  }

  getSidebarHTML(): string {
    const today         = new Date().toDateString();
    const saved         = globalContext.globalState.get<{ title: string; time: string; date: string }[]>('savedMeetings', []);
    const todayMeetings = saved.filter(m => m.date === today);
    const s             = getSettings();
    const pad           = (n: number) => String(n).padStart(2, '0');
    const dayStartVal   = `${pad(s.dayStartHour)}:${pad(s.dayStartMin)}`;
    const snoozed       = globalContext.globalState.get<SnoozedReminder[]>('snoozedReminders', []);
    const activeSnoozed = snoozed.filter(r => r.snoozeUntil > Date.now());
    const signedIn      = isSignedInToOutlook(globalContext);
    const outlookMtgs   = getCachedOutlookMeetings(globalContext);
    const lastSync      = getLastSyncTime(globalContext);

    const manualRows = todayMeetings.map((m, i) => `
      <div class="meeting-row">
        <div class="meeting-info">
          <span class="meeting-title">📅 ${escapeHtml(m.title)}</span>
          <span class="meeting-time">${escapeHtml(m.time)}</span>
        </div>
        <button class="btn-delete" onclick="deleteMeeting(${i})">🗑️</button>
      </div>`).join('');

    const outlookRows = outlookMtgs.map((m) => `
      <div class="meeting-row outlook-row">
        <div class="meeting-info">
          <span class="meeting-title">
            <span class="outlook-dot">●</span> ${escapeHtml(m.title)}
          </span>
          <span class="meeting-time">${escapeHtml(m.start)} – ${escapeHtml(m.end)}
            ${m.organizer ? `<span class="organizer">· ${escapeHtml(m.organizer)}</span>` : ''}
          </span>
          ${m.location ? `<span class="location">📍 ${escapeHtml(m.location)}</span>` : ''}
        </div>
        <div class="outlook-actions">
          ${m.joinUrl ? `<button class="btn-join" onclick="joinMeeting('${escapeHtml(m.joinUrl)}')">Join</button>` : ''}
        </div>
      </div>`).join('');

    const snoozeRows = activeSnoozed.map(r => {
      const minsLeft = Math.round((r.snoozeUntil - Date.now()) / 60000);
      return `<div class="snooze-row">😴 ${escapeHtml(r.title)} — ~${minsLeft} min</div>`;
    }).join('');

    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <style>
    *{margin:0;padding:0;box-sizing:border-box}
    body{font-family:'Segoe UI',sans-serif;background:#1e1e2e;color:#ccc;padding:12px;font-size:12px}
    .header{text-align:center;margin-bottom:14px}
    .header h2{color:#f0c040;font-size:14px;margin-bottom:6px}
    .status-badge{display:inline-block;padding:3px 12px;border-radius:20px;font-size:11px;
                  font-weight:bold;background:${isRunning ? '#2ecc71' : '#e74c3c'};color:white}
    .toggle-btns{margin-bottom:14px}
    .btn-start,.btn-stop{width:100%;padding:7px;border:none;border-radius:8px;font-size:11px;font-weight:bold;cursor:pointer}
    .btn-start{background:#2ecc71;color:white}
    .btn-stop{background:#e74c3c;color:white}
    .section-title{color:#f0c040;font-size:11px;font-weight:bold;margin-bottom:8px;
                   text-transform:uppercase;letter-spacing:1px;margin-top:14px}
    .cards{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:4px}
    .card{background:#2a2a3e;border-radius:10px;padding:10px 6px;text-align:center;border:1px solid #3a3a5e}
    .card .icon{font-size:20px}
    .card .val{color:#f0c040;font-size:11px;font-weight:bold;margin-top:3px}
    .card .lbl{color:#888;font-size:10px;margin-top:2px}

    .outlook-box{background:#2a2a3e;border-radius:10px;padding:12px;margin-bottom:4px;border:1px solid #0078d433}
    .outlook-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
    .outlook-status{font-size:10px;color:#888}
    .outlook-status.connected{color:#2ecc71}
    .btn-outlook-signin{width:100%;padding:8px;background:#0078d4;color:white;border:none;
                        border-radius:8px;font-weight:bold;font-size:12px;cursor:pointer}
    .btn-outlook-signin:hover{background:#006cbe}
    .outlook-actions-row{display:flex;gap:6px;margin-top:4px}
    .btn-sync{flex:1;padding:6px;background:#1e3a5f;color:#0078d4;border:1px solid #0078d4;
              border-radius:6px;font-size:11px;font-weight:bold;cursor:pointer}
    .btn-sync:hover{background:#0078d422}
    .btn-signout{flex:1;padding:6px;background:none;color:#e74c3c;border:1px solid #e74c3c;
                 border-radius:6px;font-size:11px;cursor:pointer}
    .btn-signout:hover{background:#e74c3c22}
    .sync-time{color:#555;font-size:10px;margin-top:5px;text-align:right}
    .outlook-row{background:#1e293b;border-radius:6px;padding:6px;margin-top:6px}
    .outlook-dot{color:#0078d4;font-size:8px;margin-right:3px}
    .organizer{color:#555;font-size:9px}
    .location{color:#888;font-size:10px;margin-top:2px;display:block}
    .btn-join{background:#0078d4;color:white;border:none;border-radius:5px;
              padding:3px 10px;font-size:10px;font-weight:bold;cursor:pointer;white-space:nowrap}
    .btn-join:hover{background:#006cbe}
    .outlook-actions{display:flex;flex-direction:column;justify-content:center;padding-left:6px}
    .no-outlook{color:#555;font-size:11px;text-align:center;padding:8px 0}

    .settings-box{background:#2a2a3e;border-radius:10px;padding:12px;margin-bottom:4px}
    .setting-row{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;gap:8px}
    .setting-row:last-child{margin-bottom:0}
    .setting-label{color:#ccc;font-size:11px;flex:1}
    .setting-row input[type=number],
    .setting-row input[type=time]{width:70px;padding:5px 7px;background:#1e1e2e;border:1px solid #444;
                     border-radius:6px;color:#f0c040;font-size:11px;outline:none;text-align:center}
    .setting-row input:focus{border-color:#f0c040}
    .setting-unit{color:#888;font-size:10px;min-width:28px}
    .btn-save{width:100%;padding:7px;background:#3498db;color:white;border:none;
              border-radius:8px;font-weight:bold;font-size:11px;cursor:pointer;margin-top:10px}
    .btn-save:hover{background:#2980b9}

    .form-section{background:#2a2a3e;border-radius:10px;padding:12px;margin-bottom:4px}
    .form-section input[type=text],
    .form-section input[type=time]{width:100%;padding:7px 9px;margin-bottom:7px;background:#1e1e2e;
                      border:1px solid #444;border-radius:7px;color:#fff;font-size:12px;outline:none}
    .form-section input:focus{border-color:#f0c040}
    .btn-add{width:100%;padding:8px;background:#f0c040;color:#1e1e2e;border:none;
             border-radius:8px;font-weight:bold;font-size:12px;cursor:pointer}
    .btn-add:hover{background:#e0b030}
    .meetings-section{background:#2a2a3e;border-radius:10px;padding:10px}
    .meeting-row{display:flex;align-items:center;justify-content:space-between;
                 padding:7px 0;border-bottom:1px solid #333}
    .meeting-row:last-child{border-bottom:none}
    .meeting-info{display:flex;flex-direction:column;gap:2px;flex:1}
    .meeting-title{color:#fff;font-size:11px}
    .meeting-time{color:#f0c040;font-size:10px}
    .btn-delete{background:none;border:none;cursor:pointer;font-size:14px;padding:2px 5px;border-radius:5px}
    .btn-delete:hover{background:#e74c3c33}
    .no-meetings{color:#666;font-size:11px;text-align:center;padding:10px 0}
    .error{color:#e74c3c;font-size:11px;margin-bottom:6px;display:none}
    .today-label{color:#888;font-size:10px;margin-bottom:8px}

    .snooze-section{background:#2a2a3e;border-radius:10px;padding:10px;margin-bottom:4px}
    .snooze-row{color:#f0c040;font-size:11px;padding:5px 0;border-bottom:1px solid #333}
    .snooze-row:last-child{border-bottom:none}
    .no-snooze{color:#666;font-size:11px;text-align:center;padding:6px 0}
  </style>
</head>
<body>

  <div class="header">
    <h2>💊 Health Reminder</h2>
    <span class="status-badge">${isRunning ? '🟢 ACTIVE' : '🔴 STOPPED'}</span>
  </div>

  <div class="toggle-btns">
    ${isRunning
      ? `<button class="btn-stop"  onclick="postMsg('stop')">■ Stop Reminders</button>`
      : `<button class="btn-start" onclick="postMsg('start')">▶ Start Reminders</button>`}
  </div>

  <div class="section-title">⏰ Active Reminders</div>
  <div class="cards">
    <div class="card"><div class="icon">☕</div><div class="val">${s.teaInterval} min</div><div class="lbl">Tea Break</div></div>
    <div class="card"><div class="icon">💧</div><div class="val">${s.waterInterval} min</div><div class="lbl">Water</div></div>
    <div class="card"><div class="icon">🧘</div><div class="val">${s.stretchInterval} min</div><div class="lbl">Stretch</div></div>
    <div class="card"><div class="icon">👁️</div><div class="val">${s.eyeInterval} min</div><div class="lbl">Eye Rest</div></div>
  </div>

  <!-- ── OUTLOOK CALENDAR ── -->
  <div class="section-title">📧 Outlook Calendar</div>
  <div class="outlook-box">
    <div class="outlook-header">
      <span style="color:#0078d4;font-size:13px;font-weight:bold">📧 Microsoft Outlook</span>
      <span class="outlook-status ${signedIn ? 'connected' : ''}">
        ${signedIn ? '🟢 Connected' : '⚪ Not connected'}
      </span>
    </div>
    ${!signedIn ? `
      <button class="btn-outlook-signin" onclick="postMsg('outlookSignIn')">
        🔑 Sign in to Outlook
      </button>
    ` : `
      <div class="outlook-actions-row">
        <button class="btn-sync"    onclick="postMsg('outlookSync')">🔄 Sync Now</button>
        <button class="btn-signout" onclick="postMsg('outlookSignOut')">Sign Out</button>
      </div>
      <div class="sync-time">Last sync: ${lastSync} · Auto every 30 min</div>
      <div style="margin-top:8px">
        ${outlookMtgs.length > 0 ? outlookRows : '<div class="no-outlook">No meetings today from Outlook</div>'}
      </div>
    `}
  </div>

  <div class="section-title">😴 Snoozed (${activeSnoozed.length})</div>
  <div class="snooze-section">
    ${activeSnoozed.length > 0 ? snoozeRows : '<div class="no-snooze">No snoozed reminders</div>'}
  </div>

  <div class="section-title">⚙️ Customize Timings</div>
  <div class="settings-box">
    <div class="setting-row">
      <span class="setting-label">🌅 Day Start Time</span>
      <input type="time" id="dayStart" value="${dayStartVal}" />
    </div>
    <div class="setting-row">
      <span class="setting-label">☕ Tea Break every</span>
      <input type="number" id="teaInterval" value="${s.teaInterval}" min="1" max="480" />
      <span class="setting-unit">min</span>
    </div>
    <div class="setting-row">
      <span class="setting-label">💧 Water every</span>
      <input type="number" id="waterInterval" value="${s.waterInterval}" min="1" max="480" />
      <span class="setting-unit">min</span>
    </div>
    <div class="setting-row">
      <span class="setting-label">🧘 Stretch every</span>
      <input type="number" id="stretchInterval" value="${s.stretchInterval}" min="1" max="480" />
      <span class="setting-unit">min</span>
    </div>
    <div class="setting-row">
      <span class="setting-label">👁️ Eye Rest every</span>
      <input type="number" id="eyeInterval" value="${s.eyeInterval}" min="1" max="480" />
      <span class="setting-unit">min</span>
    </div>
    <div class="setting-row">
      <span class="setting-label">😴 Snooze duration</span>
      <input type="number" id="snoozeMinutes" value="${s.snoozeMinutes}" min="1" max="60" />
      <span class="setting-unit">min</span>
    </div>
    <button class="btn-save" onclick="saveSettings()">💾 Save & Apply</button>
  </div>

  <div class="section-title">📅 Add Manual Meeting</div>
  <div class="form-section">
    <input type="text" id="meetingTitle" placeholder="e.g. Sprint Standup" />
    <input type="time" id="meetingTime" />
    <div class="error" id="formError">⚠️ Please fill in both fields!</div>
    <button class="btn-add" onclick="addMeeting()">+ Add Meeting</button>
  </div>

  <div class="section-title">📋 Manual Meetings (${todayMeetings.length})</div>
  <div class="today-label">📆 ${today}</div>
  <div class="meetings-section">
    ${todayMeetings.length > 0 ? manualRows : '<div class="no-meetings">No manual meetings today</div>'}
  </div>

  <script>
    const vscode = acquireVsCodeApi();
    function postMsg(cmd) { vscode.postMessage({ command: cmd }); }

    function saveSettings() {
      const dayStart = document.getElementById('dayStart').value || '09:00';
      const [dh, dm] = dayStart.split(':');
      vscode.postMessage({
        command:'saveSettings',
        teaInterval:    document.getElementById('teaInterval').value,
        waterInterval:  document.getElementById('waterInterval').value,
        stretchInterval:document.getElementById('stretchInterval').value,
        eyeInterval:    document.getElementById('eyeInterval').value,
        snoozeMinutes:  document.getElementById('snoozeMinutes').value,
        dayStartHour:   dh, dayStartMin: dm,
      });
    }

    function addMeeting() {
      const title = document.getElementById('meetingTitle').value.trim();
      const time  = document.getElementById('meetingTime').value;
      const err   = document.getElementById('formError');
      if (!title || !time) { err.style.display = 'block'; return; }
      err.style.display = 'none';
      vscode.postMessage({ command: 'addMeeting', title, time });
      document.getElementById('meetingTitle').value = '';
      document.getElementById('meetingTime').value  = '';
    }

    function deleteMeeting(i) { vscode.postMessage({ command: 'deleteMeeting', index: i }); }
    function joinMeeting(url) { vscode.postMessage({ command: 'openJoinUrl', url }); }

    document.getElementById('meetingTime').addEventListener('keydown', e => {
      if (e.key === 'Enter') { addMeeting(); }
    });
  </script>
</body>
</html>`;
  }
}

// ─── ACTIVATE ───────────────────────────────────────────────

export function activate(context: vscode.ExtensionContext) {
  globalContext = context;

  statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
  statusBarItem.text = '$(heart) Health ON';
  statusBarItem.command = 'reminder.showStatus';
  statusBarItem.show();
  context.subscriptions.push(statusBarItem);

  sidebarProvider = new HealthSidebarProvider(context);
  context.subscriptions.push(
    vscode.window.registerWebviewViewProvider(HealthSidebarProvider.viewType, sidebarProvider)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('reminder.start', () => {
      if (!isRunning) {
        startReminders(context);
        isRunning = true;
        context.globalState.update('isRunning', true);
        statusBarItem.text = '$(heart) Health ON';
        sidebarProvider.refresh();
        vscode.window.showInformationMessage('✅ Health Reminders Started!');
      }
    }),
    vscode.commands.registerCommand('reminder.stop', () => {
      stopReminders();
      isRunning = false;
      context.globalState.update('isRunning', false);
      statusBarItem.text = '$(heart) Health OFF';
      sidebarProvider.refresh();
    }),
    vscode.commands.registerCommand('reminder.addMeeting', async () => {
      const title = await vscode.window.showInputBox({ prompt: '📅 Meeting Title' });
      const time  = await vscode.window.showInputBox({ prompt: '⏰ Time HH:MM', placeHolder: '14:30' });
      if (title && time) { scheduleMeeting(context, title, time); sidebarProvider.refresh(); }
    }),
    vscode.commands.registerCommand('reminder.showStatus', () => {
      vscode.commands.executeCommand('healthReminder.sidebar.focus');
    }),
    vscode.commands.registerCommand('reminder.outlookSignIn', async () => {
      const ok = await signInToOutlook(context);
      if (ok) {
        startOutlookSync(context, (mtgs) => {
          scheduleOutlookMeetings(context, mtgs);
          sidebarProvider.refresh();
        });
        sidebarProvider.refresh();
      }
    }),
    vscode.commands.registerCommand('reminder.outlookSignOut', async () => {
      await signOutFromOutlook(context);
      sidebarProvider.refresh();
    }),
  );

  // Auto-start health reminders
  const wasRunning = context.globalState.get<boolean>('isRunning', true);
  if (wasRunning) {
    startReminders(context);
    isRunning = true;
    statusBarItem.text = '$(heart) Health ON';
  }

  loadMeetings(context);
  restoreSnoozedReminders(context);

  // Auto-resume Outlook sync if already signed in
  if (isSignedInToOutlook(context)) {
    startOutlookSync(context, (mtgs) => {
      scheduleOutlookMeetings(context, mtgs);
      sidebarProvider.refresh();
    });
  }

  sidebarProvider.refresh();

  vscode.window.showInformationMessage('💊 Health Reminder Active!', 'Open Sidebar').then(sel => {
    if (sel === 'Open Sidebar') { vscode.commands.executeCommand('healthReminder.sidebar.focus'); }
  });
}

// ─── SCHEDULE OUTLOOK MEETINGS ──────────────────────────────

function scheduleOutlookMeetings(context: vscode.ExtensionContext, mtgs: OutlookMeeting[]) {
  const s = getSettings();
  mtgs.forEach(m => {
    const [h, min] = m.start.split(':').map(Number);
    const now = new Date(), target = new Date();
    target.setHours(h, min, 0, 0);
    if (target.getTime() <= now.getTime()) { return; } // skip past meetings

    const diff = target.getTime() - now.getTime();

    if (diff > 5 * 60000) {
      setTimeout(() => {
        showPopup(context,
          `⏰ Meeting in 5 mins!`,
          `"${escapeHtml(m.title)}" starts at ${m.start}${m.joinUrl ? ' — Join link ready!' : ''}`,
          '#e67e22', s.snoozeMinutes, m.joinUrl);
      }, diff - 5 * 60000);
    }

    setTimeout(() => {
      showPopup(context,
        `📅 Meeting Now!`,
        `"${escapeHtml(m.title)}" is starting NOW!`,
        '#0078d4', s.snoozeMinutes, m.joinUrl);
    }, diff);
  });
}

// ─── RESTORE SNOOZED ────────────────────────────────────────

function restoreSnoozedReminders(context: vscode.ExtensionContext) {
  const snoozed = context.globalState.get<SnoozedReminder[]>('snoozedReminders', []);
  const now = Date.now();
  const stillPending: SnoozedReminder[] = [];

  snoozed.forEach(r => {
    const remaining = r.snoozeUntil - now;
    if (remaining > 0) {
      setTimeout(() => {
        showPopup(context, r.title, r.message, r.color, getSettings().snoozeMinutes);
        removeSnoozed(context, r.snoozeUntil);
      }, remaining);
      stillPending.push(r);
    } else {
      showPopup(context, r.title, r.message, r.color, getSettings().snoozeMinutes);
    }
  });

  context.globalState.update('snoozedReminders', stillPending);
}

function removeSnoozed(context: vscode.ExtensionContext, snoozeUntil: number) {
  const snoozed = context.globalState.get<SnoozedReminder[]>('snoozedReminders', []);
  context.globalState.update('snoozedReminders', snoozed.filter(r => r.snoozeUntil !== snoozeUntil));
}

// ─── START / STOP REMINDERS ─────────────────────────────────

function startReminders(context: vscode.ExtensionContext) {
  const s = getSettings();
  timers.forEach(t => clearInterval(t)); timers = [];

  const now = new Date();
  const dayStart = new Date();
  dayStart.setHours(s.dayStartHour, s.dayStartMin, 0, 0);
  const delayMs = now < dayStart ? dayStart.getTime() - now.getTime() : 0;

  setTimeout(() => {
    timers.push(setInterval(() => showPopup(context, '☕ Tea Break!',     'Time for a 5-min tea break!',         '#f39c12', s.snoozeMinutes), s.teaInterval     * 60000));
    timers.push(setInterval(() => showPopup(context, '💧 Drink Water!',  'Drink a glass of water now!',          '#3498db', s.snoozeMinutes), s.waterInterval   * 60000));
    timers.push(setInterval(() => showPopup(context, '🧘 Stretch Time!', 'Stand up and stretch for 2 minutes!', '#2ecc71', s.snoozeMinutes), s.stretchInterval * 60000));
    timers.push(setInterval(() => showPopup(context, '👁️ Eye Rest!',     'Look 20 feet away for 20 seconds!',   '#9b59b6', s.snoozeMinutes), s.eyeInterval     * 60000));
  }, delayMs);

  if (delayMs > 0) {
    const pad = (n: number) => String(n).padStart(2, '0');
    vscode.window.showInformationMessage(`⏰ Reminders start at ${pad(s.dayStartHour)}:${pad(s.dayStartMin)}`);
  }
}

function stopReminders() {
  timers.forEach(t => clearInterval(t)); timers = [];
  meetings.forEach(m => clearTimeout(m.timeout)); meetings = [];
  stopOutlookSync();
}

// ─── MANUAL MEETINGS ────────────────────────────────────────

function deleteMeeting(index: number, context?: vscode.ExtensionContext) {
  if (meetings[index]) { clearTimeout(meetings[index].timeout); meetings.splice(index, 1); }
  if (context) {
    const today = new Date().toDateString();
    const saved = context.globalState.get<{ title: string; time: string; date: string }[]>('savedMeetings', []);
    const todayMeetings = saved.filter(m => m.date === today);
    todayMeetings.splice(index, 1);
    context.globalState.update('savedMeetings', [...saved.filter(m => m.date !== today), ...todayMeetings]);
  }
}

function scheduleMeeting(context: vscode.ExtensionContext, title: string, time: string) {
  if (!title || title.length > 200) { vscode.window.showErrorMessage('❌ Title must be 1–200 chars'); return; }
  if (!time || !/^\d{2}:\d{2}$/.test(time)) { vscode.window.showErrorMessage('❌ Time must be HH:MM'); return; }

  const [h, m] = time.split(':').map(Number);
  if (h < 0 || h > 23 || m < 0 || m > 59) { vscode.window.showErrorMessage('❌ Invalid time'); return; }

  const now = new Date(), target = new Date();
  target.setHours(h, m, 0, 0);
  if (target.getTime() <= now.getTime()) { target.setDate(target.getDate() + 1); }

  const diff = target.getTime() - now.getTime();
  const s = getSettings();

  if (diff > 5 * 60000) {
    setTimeout(() => showPopup(context, `⏰ Meeting in 5 mins!`, `"${escapeHtml(title)}" starts soon!`, '#e67e22', s.snoozeMinutes), diff - 5 * 60000);
  }
  const t = setTimeout(() => {
    showPopup(context, `📅 Meeting Now!`, `"${escapeHtml(title)}" is starting NOW!`, '#e74c3c', s.snoozeMinutes);
  }, diff);

  meetings.push({ title, time, timeout: t });
  saveMeetings(context);
  vscode.window.showInformationMessage(`✅ "${escapeHtml(title)}" scheduled at ${time}`);
}

function saveMeetings(context: vscode.ExtensionContext) {
  const today = new Date().toDateString();
  const saved = context.globalState.get<{ title: string; time: string; date: string }[]>('savedMeetings', []);
  const todayData = meetings.map(m => ({ title: m.title, time: m.time, date: today }));
  context.globalState.update('savedMeetings', [...saved.filter(m => m.date !== today), ...todayData]);
}

function loadMeetings(context: vscode.ExtensionContext) {
  const today = new Date().toDateString();
  const saved = context.globalState.get<{ title: string; time: string; date: string }[]>('savedMeetings', []);
  const todayMeetings = saved.filter(m => m.date === today);

  meetings.forEach(m => clearTimeout(m.timeout)); meetings = [];
  const s = getSettings();

  todayMeetings.forEach(m => {
    const [h, min] = m.time.split(':').map(Number);
    const now = new Date(), target = new Date();
    target.setHours(h, min, 0, 0);

    if (target.getTime() > now.getTime()) {
      const diff = target.getTime() - now.getTime();
      if (diff > 5 * 60000) {
        setTimeout(() => showPopup(context, `⏰ Meeting in 5 mins!`, `"${escapeHtml(m.title)}" starts soon!`, '#e67e22', s.snoozeMinutes), diff - 5 * 60000);
      }
      const t = setTimeout(() => {
        showPopup(context, `📅 Meeting Now!`, `"${escapeHtml(m.title)}" is starting NOW!`, '#e74c3c', s.snoozeMinutes);
      }, diff);
      meetings.push({ title: m.title, time: m.time, timeout: t });
    } else {
      meetings.push({ title: m.title, time: m.time, timeout: setTimeout(() => {}, 0) });
    }
  });
}

// ─── SHOW POPUP ─────────────────────────────────────────────

function showPopup(
  context: vscode.ExtensionContext,
  title: string, message: string, color: string,
  snoozeMinutes: number, joinUrl?: string
) {
  const panel = vscode.window.createWebviewPanel(
    'healthReminder', title, vscode.ViewColumn.One, { enableScripts: true }
  );
  panel.webview.html = getPopupHTML(title, message, color, snoozeMinutes, joinUrl);
  const autoClose = setTimeout(() => { try { panel.dispose(); } catch (e) {} }, 30000);

  panel.webview.onDidReceiveMessage(msg => {
    if (msg.command === 'snooze') {
      clearTimeout(autoClose);
      try { panel.dispose(); } catch (e) {}
      const snoozeUntil = Date.now() + snoozeMinutes * 60000;
      const snoozed = context.globalState.get<SnoozedReminder[]>('snoozedReminders', []);
      snoozed.push({ title, message, color, snoozeUntil });
      context.globalState.update('snoozedReminders', snoozed);
      setTimeout(() => {
        showPopup(context, title, message, color, snoozeMinutes, joinUrl);
        removeSnoozed(context, snoozeUntil);
        sidebarProvider.refresh();
      }, snoozeMinutes * 60000);
      vscode.window.showInformationMessage(`😴 Snoozed! "${title}" returns in ${snoozeMinutes} min`);
      sidebarProvider.refresh();
    }
    if (msg.command === 'close') {
      clearTimeout(autoClose); try { panel.dispose(); } catch (e) {}
    }
    if (msg.command === 'join' && joinUrl) {
      vscode.env.openExternal(vscode.Uri.parse(joinUrl));
    }
  });
}

// ─── POPUP HTML ─────────────────────────────────────────────

function getPopupHTML(title: string, message: string, color: string, snoozeMinutes: number, joinUrl?: string): string {
  const sc = sanitizeColor(color);
  const et = escapeHtml(title);
  const em = escapeHtml(message);
  const emoji = title.includes('Tea') ? '☕' : title.includes('Water') ? '💧'
    : title.includes('Stretch') ? '🧘' : title.includes('Eye') ? '👁️' : '📅';

  return `<!DOCTYPE html>
<html>
<head>
  <style>
    *{margin:0;padding:0;box-sizing:border-box}
    body{font-family:'Segoe UI',sans-serif;background:#1e1e2e;display:flex;
         justify-content:center;align-items:center;min-height:100vh}
    .box{background:#2a2a3e;border-radius:20px;padding:40px;text-align:center;
         max-width:480px;width:90%;border:2px solid ${sc};position:relative;
         box-shadow:0 0 40px ${sc}44;animation:pop 0.4s ease}
    @keyframes pop{from{transform:scale(0.5);opacity:0}to{transform:scale(1);opacity:1}}
    .close-btn{position:absolute;top:14px;right:16px;background:none;border:none;
               color:#888;font-size:22px;cursor:pointer;padding:4px 8px;border-radius:6px}
    .close-btn:hover{background:#ffffff22;color:#fff}
    .icon{font-size:72px;margin-bottom:16px;display:block;animation:bounce 1s infinite}
    @keyframes bounce{0%,100%{transform:translateY(0)}50%{transform:translateY(-10px)}}
    h1{color:${sc};font-size:1.8em;margin-bottom:12px}
    p{color:#ccc;font-size:1.1em;margin-bottom:24px}
    .btns{display:flex;gap:12px;justify-content:center;flex-wrap:wrap}
    button{padding:10px 24px;border:none;border-radius:10px;font-size:1em;
           font-weight:bold;cursor:pointer;transition:transform 0.2s}
    button:hover{transform:scale(1.05)}
    .stop{background:${sc};color:#fff}
    .done{background:#2ecc71;color:#fff}
    .snooze{background:#555;color:#fff}
    .join{background:#0078d4;color:#fff}
    .bar{width:100%;height:5px;background:#333;border-radius:3px;margin-top:20px;overflow:hidden}
    .fill{height:100%;background:${sc};animation:shrink 30s linear forwards}
    @keyframes shrink{from{width:100%}to{width:0}}
    .countdown{color:#888;font-size:11px;margin-top:10px}
    .snooze-info{color:${sc};font-size:10px;margin-top:5px;opacity:0.8}
  </style>
</head>
<body>
  <div class="box">
    <button class="close-btn" onclick="closeAll()">✕</button>
    <span class="icon">${emoji}</span>
    <h1>${et}</h1>
    <p>${em}</p>
    <div class="btns">
      <button class="stop"   onclick="stopAlarm()">🔇 Stop</button>
      <button class="snooze" onclick="snooze()">😴 Snooze ${snoozeMinutes}min</button>
      <button class="done"   onclick="closeAll()">✅ Done</button>
      ${joinUrl ? `<button class="join" onclick="joinMeeting()">🔗 Join Meeting</button>` : ''}
    </div>
    <div class="bar"><div class="fill"></div></div>
    <div class="countdown" id="cd">Auto-closing in 30 seconds...</div>
    <div class="snooze-info">😴 Snooze returns in ${snoozeMinutes} minutes</div>
  </div>
  <script>
    const vscode = acquireVsCodeApi();
    let ac=null,iv=null,cdInterval=null;
    function initAudio(){if(!ac){ac=new(window.AudioContext||window.webkitAudioContext)();}if(ac.state==='suspended'){ac.resume();}}
    const notes=[523,659,784,1047,784,659,523];
    function playMelody(){initAudio();const t=ac.currentTime;notes.forEach((freq,i)=>{const o=ac.createOscillator(),g=ac.createGain();o.connect(g);g.connect(ac.destination);o.type='triangle';o.frequency.value=freq;g.gain.setValueAtTime(0,t+i*0.18);g.gain.linearRampToValueAtTime(0.4,t+i*0.18+0.05);g.gain.exponentialRampToValueAtTime(0.001,t+i*0.18+0.16);o.start(t+i*0.18);o.stop(t+i*0.18+0.18);});}
    function startAlarm(){try{initAudio();playMelody();iv=setInterval(playMelody,3000);}catch(e){}}
    startAlarm();
    document.addEventListener('click',function onFC(){if(!iv){startAlarm();}document.removeEventListener('click',onFC);},{once:true});
    let secs=30;const cd=document.getElementById('cd');
    cdInterval=setInterval(()=>{secs--;if(cd){cd.textContent='Auto-closing in '+secs+' second'+(secs!==1?'s':'')+'...';}if(secs<=0){clearInterval(cdInterval);}},1000);
    function stopAlarm(){if(iv){clearInterval(iv);iv=null;}if(ac){ac.suspend();}if(cd){cd.textContent='🔇 Alarm stopped';}}
    function snooze(){if(iv){clearInterval(iv);iv=null;}if(ac){ac.suspend();}clearInterval(cdInterval);if(cd){cd.textContent='😴 Snoozed! Returns in ${snoozeMinutes} min...';}vscode.postMessage({command:'snooze'});}
    function closeAll(){if(iv){clearInterval(iv);iv=null;}if(cdInterval){clearInterval(cdInterval);cdInterval=null;}if(ac){ac.suspend();}vscode.postMessage({command:'close'});}
    function joinMeeting(){vscode.postMessage({command:'join'});}
  </script>
</body>
</html>`;
}

export function deactivate() { stopReminders(); }