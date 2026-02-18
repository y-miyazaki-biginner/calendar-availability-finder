// app.js - Calendar Availability Finder (Web App version)
// Google Identity Services (GIS) を使ったブラウザOAuth2フロー

// =====================================================
// ★ デプロイ時にここを書き換える
// =====================================================
const CONFIG = {
  CLIENT_ID: 'YOUR_CLIENT_ID.apps.googleusercontent.com',
  SCOPES: 'https://www.googleapis.com/auth/calendar.readonly https://www.googleapis.com/auth/calendar.freebusy',
};

class CalendarAvailabilityFinder {
  constructor() {
    this.token = null;
    this.tokenClient = null;
    this.emails = [];
    this.settings = {
      searchRange: 14,
      startTime: '11:00',
      endTime: '18:00',
      meetingDuration: 30,
      activeDays: [1, 2, 3, 4, 5],
      excludeKeywords: ['画面操作'],
    };

    this.init();
  }

  async init() {
    this.bindElements();
    this.bindEvents();
    this.loadSettings();
    this.initGoogleAuth();
  }

  // ============================================================
  // Google Identity Services (Web OAuth2)
  // ============================================================
  initGoogleAuth() {
    // GIS ライブラリがロードされるまで待機
    const waitForGis = () => {
      if (typeof google !== 'undefined' && google.accounts) {
        this.tokenClient = google.accounts.oauth2.initTokenClient({
          client_id: CONFIG.CLIENT_ID,
          scope: CONFIG.SCOPES,
          callback: (response) => {
            if (response.error) {
              console.error('OAuth error:', response);
              return;
            }
            this.token = response.access_token;
            this.showMain();
          },
        });
      } else {
        setTimeout(waitForGis, 100);
      }
    };
    waitForGis();
  }

  // ============================================================
  // Elements & Events
  // ============================================================
  bindElements() {
    this.authSection = document.getElementById('auth-section');
    this.mainSection = document.getElementById('main-section');
    this.authBtn = document.getElementById('auth-btn');
    this.logoutBtn = document.getElementById('logout-btn');
    this.settingsToggle = document.getElementById('settings-toggle');
    this.settingsPanel = document.getElementById('settings-panel');
    this.searchRange = document.getElementById('search-range');
    this.timeStart = document.getElementById('time-start');
    this.timeEnd = document.getElementById('time-end');
    this.meetingDuration = document.getElementById('meeting-duration');
    this.excludeKeywords = document.getElementById('exclude-keywords');
    this.saveSettingsBtn = document.getElementById('save-settings');
    this.emailInput = document.getElementById('email-input');
    this.addEmailBtn = document.getElementById('add-email-btn');
    this.emailTags = document.getElementById('email-tags');
    this.searchBtn = document.getElementById('search-btn');
    this.loading = document.getElementById('loading');
    this.results = document.getElementById('results');
    this.freeSection = document.getElementById('free-section');
    this.freeSlots = document.getElementById('free-slots');
    this.partialSection = document.getElementById('partial-section');
    this.partialSlots = document.getElementById('partial-slots');
    this.conflictsSection = document.getElementById('conflicts-section');
    this.conflictsList = document.getElementById('conflicts-list');
  }

  bindEvents() {
    this.authBtn.addEventListener('click', () => this.authenticate());
    this.logoutBtn.addEventListener('click', () => this.logout());
    this.settingsToggle.addEventListener('click', () => this.toggleSettings());
    this.saveSettingsBtn.addEventListener('click', () => this.saveSettings());
    this.addEmailBtn.addEventListener('click', () => this.addEmail());
    this.emailInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') this.addEmail();
    });
    this.searchBtn.addEventListener('click', () => this.searchAvailability());
  }

  // ============================================================
  // Auth
  // ============================================================
  authenticate() {
    if (this.tokenClient) {
      this.tokenClient.requestAccessToken();
    }
  }

  logout() {
    if (this.token) {
      google.accounts.oauth2.revoke(this.token);
    }
    this.token = null;
    this.emails = [];
    this.showAuth();
  }

  showAuth() {
    this.authSection.classList.remove('hidden');
    this.mainSection.classList.add('hidden');
  }

  showMain() {
    this.authSection.classList.add('hidden');
    this.mainSection.classList.remove('hidden');
    this.applySettingsToUI();
  }

  // ============================================================
  // Settings (localStorage for persistence)
  // ============================================================
  loadSettings() {
    try {
      const saved = localStorage.getItem('calendarSettings');
      if (saved) {
        this.settings = { ...this.settings, ...JSON.parse(saved) };
      }
    } catch { /* ignore */ }
  }

  saveSettings() {
    this.settings.searchRange = parseInt(this.searchRange.value);
    this.settings.startTime = this.timeStart.value;
    this.settings.endTime = this.timeEnd.value;
    this.settings.meetingDuration = parseInt(this.meetingDuration.value);

    // 除外キーワード
    const rawKeywords = this.excludeKeywords.value;
    this.settings.excludeKeywords = rawKeywords
      .split(/[,、，]/)
      .map((k) => k.trim())
      .filter(Boolean);

    // 対象曜日
    const dayCheckboxes = document.querySelectorAll('.day-check input');
    this.settings.activeDays = [];
    dayCheckboxes.forEach((cb) => {
      if (cb.checked) this.settings.activeDays.push(parseInt(cb.value));
    });

    localStorage.setItem('calendarSettings', JSON.stringify(this.settings));
    this.saveSettingsBtn.textContent = '保存しました!';
    setTimeout(() => {
      this.saveSettingsBtn.textContent = '設定を保存';
    }, 1500);
  }

  applySettingsToUI() {
    this.searchRange.value = this.settings.searchRange;
    this.timeStart.value = this.settings.startTime;
    this.timeEnd.value = this.settings.endTime;
    this.meetingDuration.value = this.settings.meetingDuration;
    this.excludeKeywords.value = (this.settings.excludeKeywords || []).join(', ');

    const dayCheckboxes = document.querySelectorAll('.day-check input');
    dayCheckboxes.forEach((cb) => {
      cb.checked = this.settings.activeDays.includes(parseInt(cb.value));
    });
  }

  toggleSettings() {
    this.settingsPanel.classList.toggle('hidden');
  }

  // ============================================================
  // Email Management
  // ============================================================
  addEmail() {
    const email = this.emailInput.value.trim().toLowerCase();
    if (!email) return;
    if (!this.isValidEmail(email)) {
      this.emailInput.style.borderColor = '#d93025';
      setTimeout(() => { this.emailInput.style.borderColor = ''; }, 2000);
      return;
    }
    if (this.emails.includes(email)) {
      this.emailInput.value = '';
      return;
    }
    this.emails.push(email);
    this.renderEmailTags();
    this.emailInput.value = '';
    this.emailInput.focus();
    this.updateSearchButton();
  }

  removeEmail(email) {
    this.emails = this.emails.filter((e) => e !== email);
    this.renderEmailTags();
    this.updateSearchButton();
  }

  renderEmailTags() {
    this.emailTags.innerHTML = this.emails
      .map(
        (email) => `
        <span class="email-tag">
          ${this.escapeHtml(email)}
          <span class="remove-email" data-email="${this.escapeHtml(email)}">&times;</span>
        </span>`
      )
      .join('');
    this.emailTags.querySelectorAll('.remove-email').forEach((btn) => {
      btn.addEventListener('click', () => this.removeEmail(btn.dataset.email));
    });
  }

  updateSearchButton() {
    this.searchBtn.disabled = this.emails.length === 0;
  }

  isValidEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }

  // ============================================================
  // ① 除外キーワードフィルタ
  // ============================================================
  shouldExcludeEvent(title) {
    if (!title) return false;
    const keywords = this.settings.excludeKeywords || [];
    const lowerTitle = title.toLowerCase();
    return keywords.some((kw) => kw && lowerTitle.includes(kw.toLowerCase()));
  }

  // ============================================================
  // ② 検索時間帯内かどうか判定
  // ============================================================
  isWithinSearchTimeRange(eventStart, eventEnd) {
    const [startH, startM] = this.settings.startTime.split(':').map(Number);
    const [endH, endM] = this.settings.endTime.split(':').map(Number);
    const searchStartMin = startH * 60 + startM;
    const searchEndMin = endH * 60 + endM;

    const eventStartMin = eventStart.getHours() * 60 + eventStart.getMinutes();
    const eventEndMin = eventEnd.getHours() * 60 + eventEnd.getMinutes();

    // 検索時間帯と重なるかチェック（日付は別で見るので時刻部分のみ）
    return eventStartMin < searchEndMin && eventEndMin > searchStartMin;
  }

  // ============================================================
  // Google Calendar API (直接 fetch)
  // ============================================================
  async apiQueryFreeBusy(params) {
    const res = await fetch('https://www.googleapis.com/calendar/v3/freeBusy', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${this.token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(params),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `API error: ${res.status}`);
    }
    return res.json();
  }

  async apiGetEvents(calendarId, timeMin, timeMax) {
    const params = new URLSearchParams({
      timeMin, timeMax,
      singleEvents: 'true',
      orderBy: 'startTime',
      maxResults: '250',
    });
    const res = await fetch(
      `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(calendarId)}/events?${params}`,
      { headers: { 'Authorization': `Bearer ${this.token}` } }
    );
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `API error: ${res.status}`);
    }
    return res.json();
  }

  // ============================================================
  // Main Search
  // ============================================================
  async searchAvailability() {
    if (this.emails.length === 0) return;

    this.showLoading(true);
    this.results.classList.add('hidden');
    this.conflictsSection.classList.add('hidden');
    this.partialSection.classList.add('hidden');

    try {
      const now = new Date();
      const timeMin = new Date(now);
      timeMin.setMinutes(0, 0, 0);
      timeMin.setHours(timeMin.getHours() + 1);

      const timeMax = new Date(now);
      timeMax.setDate(timeMax.getDate() + this.settings.searchRange);
      timeMax.setHours(23, 59, 59, 0);

      // --- 全参加者のイベント詳細を取得 ---
      const allEvents = await this.getAllEvents(timeMin, timeMax);

      // --- ①除外キーワードフィルタ適用 ---
      const filteredEvents = allEvents.filter(
        (ev) => !this.shouldExcludeEvent(ev.title)
      );

      // --- busyPeriods を filteredEvents から構築 ---
      const busyPeriods = filteredEvents.map((ev) => ({
        email: ev.email,
        start: ev.start,
        end: ev.end,
      }));

      // --- ③ 競合なし + 競合少スロットを算出 ---
      const { freeSlots, partialSlots } = this.findAllSlots(
        timeMin, timeMax, busyPeriods
      );

      // --- ② 検索時間帯内の競合のみ抽出 ---
      const conflictsInRange = filteredEvents.filter((ev) =>
        this.isWithinSearchTimeRange(ev.start, ev.end) &&
        this.settings.activeDays.includes(ev.start.getDay())
      );

      this.renderResults(freeSlots, partialSlots, conflictsInRange);
    } catch (error) {
      this.showError(error.message);
    } finally {
      this.showLoading(false);
    }
  }

  async getAllEvents(timeMin, timeMax) {
    const allEvents = [];

    for (const email of this.emails) {
      try {
        const data = await this.apiGetEvents(
          email, timeMin.toISOString(), timeMax.toISOString()
        );
        if (data.items) {
          for (const event of data.items) {
            if (event.start?.dateTime) {
              allEvents.push({
                email,
                title: event.summary || '(タイトルなし)',
                start: new Date(event.start.dateTime),
                end: new Date(event.end.dateTime),
              });
            }
          }
        }
      } catch {
        // FreeBusy API へフォールバック
        try {
          const fbData = await this.apiQueryFreeBusy({
            timeMin: timeMin.toISOString(),
            timeMax: timeMax.toISOString(),
            timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
            items: [{ id: email }],
          });
          const cal = fbData.calendars?.[email];
          if (cal?.busy) {
            for (const b of cal.busy) {
              allEvents.push({
                email,
                title: '(詳細不明 - FreeBusy)',
                start: new Date(b.start),
                end: new Date(b.end),
              });
            }
          }
        } catch { /* skip */ }
      }
    }

    return allEvents;
  }

  // ============================================================
  // ③ スロット算出: 競合なし & 競合少
  // ============================================================
  findAllSlots(timeMin, timeMax, busyPeriods) {
    const freeSlots = [];
    const partialSlots = [];
    const [startHour, startMin] = this.settings.startTime.split(':').map(Number);
    const [endHour, endMin] = this.settings.endTime.split(':').map(Number);
    const duration = this.settings.meetingDuration;
    const slotStep = 30;
    const totalPeople = this.emails.length;

    const current = new Date(timeMin);
    current.setHours(startHour, startMin, 0, 0);

    if (current < timeMin) {
      current.setTime(timeMin.getTime());
      current.setHours(startHour, startMin, 0, 0);
      if (current < timeMin) {
        current.setDate(current.getDate() + 1);
        current.setHours(startHour, startMin, 0, 0);
      }
    }

    const endTimeMinutes = endHour * 60 + endMin;

    while (current < timeMax) {
      const dayOfWeek = current.getDay();

      if (!this.settings.activeDays.includes(dayOfWeek)) {
        current.setDate(current.getDate() + 1);
        current.setHours(startHour, startMin, 0, 0);
        continue;
      }

      const slotStart = new Date(current);
      const slotEnd = new Date(current);
      slotEnd.setMinutes(slotEnd.getMinutes() + duration);

      const slotEndMinutes = slotEnd.getHours() * 60 + slotEnd.getMinutes();

      if (slotEndMinutes > endTimeMinutes || slotEnd.getDate() !== slotStart.getDate()) {
        current.setDate(current.getDate() + 1);
        current.setHours(startHour, startMin, 0, 0);
        continue;
      }

      if (slotStart > new Date()) {
        // このスロットと競合する予定を集める
        const conflicting = busyPeriods.filter(
          (busy) => slotStart < busy.end && slotEnd > busy.start
        );

        // 競合している人数（ユニーク）
        const conflictingPeople = new Set(conflicting.map((c) => c.email));
        const conflictCount = conflictingPeople.size;

        if (conflictCount === 0) {
          freeSlots.push({
            start: new Date(slotStart),
            end: new Date(slotEnd),
            conflictCount: 0,
            conflictingEmails: [],
          });
        } else if (conflictCount < totalPeople) {
          // 全員がNGではない = 一部の人だけ競合
          partialSlots.push({
            start: new Date(slotStart),
            end: new Date(slotEnd),
            conflictCount,
            conflictingEmails: [...conflictingPeople],
            conflictingEvents: conflicting,
          });
        }
        // conflictCount === totalPeople → 全員NG → 表示しない
      }

      current.setMinutes(current.getMinutes() + slotStep);
    }

    // 競合少ない順にソート
    partialSlots.sort((a, b) => {
      if (a.conflictCount !== b.conflictCount) return a.conflictCount - b.conflictCount;
      return a.start - b.start;
    });

    return { freeSlots, partialSlots };
  }

  // ============================================================
  // Rendering
  // ============================================================
  renderResults(freeSlots, partialSlots, conflictsInRange) {
    this.results.classList.remove('hidden');

    // --- 競合なしスロット ---
    if (freeSlots.length === 0) {
      this.freeSlots.innerHTML = `
        <div class="no-results">
          <p>競合なしの空き時間は見つかりませんでした。</p>
          <p style="font-size: 12px; margin-top: 4px;">下の「競合あり」セクションを確認してください。</p>
        </div>`;
    } else {
      this.freeSlots.innerHTML = this.renderSlotCards(freeSlots, 'free');
      this.attachCopyButtons(this.freeSlots);
    }

    // --- 競合少スロット ---
    if (partialSlots.length > 0) {
      this.partialSection.classList.remove('hidden');
      this.partialSlots.innerHTML = this.renderSlotCards(partialSlots, 'partial');
      this.attachCopyButtons(this.partialSlots);
    } else {
      this.partialSection.classList.add('hidden');
    }

    // --- ② 検索時間帯内の競合のみ表示 ---
    if (conflictsInRange.length > 0) {
      this.conflictsSection.classList.remove('hidden');

      const byPerson = {};
      for (const c of conflictsInRange) {
        if (!byPerson[c.email]) byPerson[c.email] = [];
        byPerson[c.email].push(c);
      }

      let html = '';
      for (const [email, events] of Object.entries(byPerson)) {
        html += `<div style="font-size: 13px; font-weight: 600; margin: 8px 0 4px; color: #202124;">${this.escapeHtml(email)}</div>`;
        for (const event of events) {
          html += `
            <div class="conflict-card">
              <div class="conflict-title">${this.escapeHtml(event.title)}</div>
              <div class="conflict-time">${this.formatDate(event.start)} ${this.formatTimeRange(event.start, event.end)}</div>
            </div>`;
        }
      }
      this.conflictsList.innerHTML = html;
    } else {
      this.conflictsSection.classList.add('hidden');
    }
  }

  renderSlotCards(slots, type) {
    const grouped = this.groupSlotsByDate(slots);
    let html = '';

    for (const [dateStr, daySlots] of Object.entries(grouped)) {
      html += `<div class="slot-date">${dateStr}</div>`;
      for (const slot of daySlots) {
        const timeStr = this.formatTimeRange(slot.start, slot.end);
        const copyText = `${dateStr} ${timeStr}`;

        let badge = '';
        if (type === 'partial') {
          const names = slot.conflictingEmails.map((e) => e.split('@')[0]).join(', ');
          badge = `<span class="slot-conflict-badge">${slot.conflictCount}人競合 (${this.escapeHtml(names)})</span>`;
        }

        html += `
          <div class="slot-card slot-card-${type}">
            <div class="slot-header">
              <div>
                <span class="slot-time">${timeStr}</span>
                ${badge}
              </div>
              <button class="btn btn-copy" data-copy="${this.escapeHtml(copyText)}">コピー</button>
            </div>
          </div>`;
      }
    }

    return html;
  }

  attachCopyButtons(container) {
    container.querySelectorAll('.btn-copy').forEach((btn) => {
      btn.addEventListener('click', () => {
        navigator.clipboard.writeText(btn.dataset.copy);
        btn.textContent = 'コピー済み';
        setTimeout(() => { btn.textContent = 'コピー'; }, 1500);
      });
    });
  }

  showError(message) {
    this.results.classList.remove('hidden');
    this.freeSlots.innerHTML = `<div class="error-message">${this.escapeHtml(message)}</div>`;
  }

  showLoading(show) {
    if (show) {
      this.loading.classList.remove('hidden');
      this.searchBtn.disabled = true;
    } else {
      this.loading.classList.add('hidden');
      this.searchBtn.disabled = this.emails.length === 0;
    }
  }

  // ============================================================
  // Helpers
  // ============================================================
  groupSlotsByDate(slots) {
    const grouped = {};
    for (const slot of slots) {
      const dateStr = this.formatDate(slot.start);
      if (!grouped[dateStr]) grouped[dateStr] = [];
      grouped[dateStr].push(slot);
    }
    return grouped;
  }

  formatDate(date) {
    const days = ['日', '月', '火', '水', '木', '金', '土'];
    return `${date.getMonth() + 1}/${date.getDate()}(${days[date.getDay()]})`;
  }

  formatTimeRange(start, end) {
    const fmt = (d) =>
      `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
    return `${fmt(start)} - ${fmt(end)}`;
  }

  escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  new CalendarAvailabilityFinder();
});
