// app.js - Calendar Availability Finder (Web App version)
// Google Identity Services (GIS) を使ったブラウザOAuth2フロー

// =====================================================
// ★ デプロイ時にここを書き換える
// =====================================================
const CONFIG = {
  CLIENT_ID: '416943777269-ie7jg6j4tr53j1lqfplvcnhde0rajuls.apps.googleusercontent.com',
  SCOPES: 'https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.freebusy',
};

class CalendarAvailabilityFinder {
  constructor() {
    this.token = null;
    this.tokenClient = null;
    this.emails = [];
    this.settings = {
      rangeMode: 'relative',   // 'relative' or 'absolute'
      searchRange: 14,
      dateStart: '',
      dateEnd: '',
      startTime: '11:00',
      endTime: '18:00',
      meetingDuration: 30,
      activeDays: [1, 2, 3, 4, 5],
      excludeKeywords: ['画面操作'],
    };
    this.lastFreeSlots = [];
    this.lastPartialSlots = [];

    this.init();
  }

  async init() {
    this.bindElements();
    this.bindEvents();
    this.loadSettings();
    this.loadEmailHistory();
    this.loadSavedGroups();
    this.initGoogleAuth();
  }

  // ============================================================
  // Google Identity Services (Web OAuth2)
  // ① ログイン保持: token を sessionStorage に保存し再訪問時に復元
  // ============================================================
  initGoogleAuth() {
    // sessionStorage から token を復元
    const savedToken = sessionStorage.getItem('calendarToken');
    if (savedToken) {
      this.token = savedToken;
      // token が有効か確認
      this.validateToken().then((valid) => {
        if (valid) {
          this.showMain();
        } else {
          sessionStorage.removeItem('calendarToken');
          this.token = null;
        }
      });
    }

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
            // sessionStorage に保存（タブ閉じるまで有効）
            sessionStorage.setItem('calendarToken', this.token);
            this.showMain();
          },
        });
      } else {
        setTimeout(waitForGis, 100);
      }
    };
    waitForGis();
  }

  async validateToken() {
    try {
      const res = await fetch(
        `https://www.googleapis.com/oauth2/v1/tokeninfo?access_token=${this.token}`
      );
      return res.ok;
    } catch {
      return false;
    }
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
    this.rangeMode = document.getElementById('range-mode');
    this.relativeRangeGroup = document.getElementById('relative-range-group');
    this.absoluteRangeGroup = document.getElementById('absolute-range-group');
    this.searchRange = document.getElementById('search-range');
    this.dateStart = document.getElementById('date-start');
    this.dateEnd = document.getElementById('date-end');
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
    this.registerPanel = document.getElementById('register-panel');
    this.eventTitle = document.getElementById('event-title');
    this.registerBtn = document.getElementById('register-btn');
    this.registerSummary = document.getElementById('register-summary');
    this.registerStatus = document.getElementById('register-status');
    // メールアドレス保存
    this.savedGroupsContainer = document.getElementById('saved-groups');
    this.saveGroupBtn = document.getElementById('save-group-btn');
    this.emailHistoryList = document.getElementById('email-history-list');
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
    this.registerBtn.addEventListener('click', () => this.registerEvents());
    this.eventTitle.addEventListener('input', () => this.updateRegisterButton());
    // 範囲モード切替
    this.rangeMode.addEventListener('change', () => this.onRangeModeChange());
    // グループ保存
    this.saveGroupBtn.addEventListener('click', () => this.saveCurrentGroup());
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
    sessionStorage.removeItem('calendarToken');
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
    this.renderSavedGroups();
  }

  // ============================================================
  // Settings (localStorage)
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
    this.settings.rangeMode = this.rangeMode.value;
    this.settings.searchRange = parseInt(this.searchRange.value);
    this.settings.dateStart = this.dateStart.value;
    this.settings.dateEnd = this.dateEnd.value;
    this.settings.startTime = this.timeStart.value;
    this.settings.endTime = this.timeEnd.value;
    this.settings.meetingDuration = parseInt(this.meetingDuration.value);

    const rawKeywords = this.excludeKeywords.value;
    this.settings.excludeKeywords = rawKeywords
      .split(/[,、，]/)
      .map((k) => k.trim())
      .filter(Boolean);

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
    this.rangeMode.value = this.settings.rangeMode || 'relative';
    this.searchRange.value = this.settings.searchRange;
    this.timeStart.value = this.settings.startTime;
    this.timeEnd.value = this.settings.endTime;
    this.meetingDuration.value = this.settings.meetingDuration;
    this.excludeKeywords.value = (this.settings.excludeKeywords || []).join(', ');

    // 日付指定のデフォルト値を設定
    if (!this.settings.dateStart) {
      const today = new Date();
      this.dateStart.value = this.toDateString(today);
      const twoWeeks = new Date(today);
      twoWeeks.setDate(twoWeeks.getDate() + 14);
      this.dateEnd.value = this.toDateString(twoWeeks);
    } else {
      this.dateStart.value = this.settings.dateStart;
      this.dateEnd.value = this.settings.dateEnd;
    }

    const dayCheckboxes = document.querySelectorAll('.day-check input');
    dayCheckboxes.forEach((cb) => {
      cb.checked = this.settings.activeDays.includes(parseInt(cb.value));
    });

    this.onRangeModeChange();
  }

  onRangeModeChange() {
    const mode = this.rangeMode.value;
    if (mode === 'relative') {
      this.relativeRangeGroup.classList.remove('hidden');
      this.absoluteRangeGroup.classList.add('hidden');
    } else {
      this.relativeRangeGroup.classList.add('hidden');
      this.absoluteRangeGroup.classList.remove('hidden');
    }
  }

  toggleSettings() {
    this.settingsPanel.classList.toggle('hidden');
  }

  toDateString(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // ============================================================
  // ④ Email Management (履歴保存 + グループ保存)
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
    this.addToEmailHistory(email);
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

  // メール履歴（datalist用）
  loadEmailHistory() {
    try {
      const history = JSON.parse(localStorage.getItem('emailHistory') || '[]');
      this.emailHistory = history;
      this.renderEmailHistory();
    } catch {
      this.emailHistory = [];
    }
  }

  addToEmailHistory(email) {
    if (!this.emailHistory.includes(email)) {
      this.emailHistory.push(email);
      localStorage.setItem('emailHistory', JSON.stringify(this.emailHistory));
      this.renderEmailHistory();
    }
  }

  renderEmailHistory() {
    this.emailHistoryList.innerHTML = this.emailHistory
      .map((e) => `<option value="${this.escapeHtml(e)}">`)
      .join('');
  }

  // グループ保存
  loadSavedGroups() {
    try {
      this.savedGroups = JSON.parse(localStorage.getItem('emailGroups') || '[]');
    } catch {
      this.savedGroups = [];
    }
  }

  saveCurrentGroup() {
    if (this.emails.length === 0) return;
    const name = prompt('グループ名を入力してください:', `グループ${this.savedGroups.length + 1}`);
    if (!name) return;
    this.savedGroups.push({ name, emails: [...this.emails] });
    localStorage.setItem('emailGroups', JSON.stringify(this.savedGroups));
    this.renderSavedGroups();
  }

  loadGroup(index) {
    const group = this.savedGroups[index];
    if (!group) return;
    this.emails = [...group.emails];
    this.renderEmailTags();
    this.updateSearchButton();
  }

  deleteGroup(index, e) {
    e.stopPropagation();
    this.savedGroups.splice(index, 1);
    localStorage.setItem('emailGroups', JSON.stringify(this.savedGroups));
    this.renderSavedGroups();
  }

  renderSavedGroups() {
    if (this.savedGroups.length === 0) {
      this.savedGroupsContainer.innerHTML = '';
      return;
    }
    this.savedGroupsContainer.innerHTML = this.savedGroups
      .map((g, i) => `
        <span class="group-tag" data-index="${i}" title="${g.emails.join(', ')}">
          ${this.escapeHtml(g.name)} (${g.emails.length}人)
          <span class="group-delete" data-index="${i}">&times;</span>
        </span>`)
      .join('');

    this.savedGroupsContainer.querySelectorAll('.group-tag').forEach((tag) => {
      tag.addEventListener('click', (e) => {
        if (!e.target.classList.contains('group-delete')) {
          this.loadGroup(parseInt(tag.dataset.index));
        }
      });
    });
    this.savedGroupsContainer.querySelectorAll('.group-delete').forEach((btn) => {
      btn.addEventListener('click', (e) => this.deleteGroup(parseInt(btn.dataset.index), e));
    });
  }

  // ============================================================
  // 除外キーワードフィルタ
  // ============================================================
  shouldExcludeEvent(title) {
    if (!title) return false;
    const keywords = this.settings.excludeKeywords || [];
    const lowerTitle = title.toLowerCase();
    return keywords.some((kw) => kw && lowerTitle.includes(kw.toLowerCase()));
  }

  // ============================================================
  // 検索時間帯内かどうか判定
  // ============================================================
  isWithinSearchTimeRange(eventStart, eventEnd) {
    const [startH, startM] = this.settings.startTime.split(':').map(Number);
    const [endH, endM] = this.settings.endTime.split(':').map(Number);
    const searchStartMin = startH * 60 + startM;
    const searchEndMin = endH * 60 + endM;

    const eventStartMin = eventStart.getHours() * 60 + eventStart.getMinutes();
    const eventEndMin = eventEnd.getHours() * 60 + eventEnd.getMinutes();

    return eventStartMin < searchEndMin && eventEndMin > searchStartMin;
  }

  // ============================================================
  // Google Calendar API
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

  async apiInsertEvent(eventBody) {
    const res = await fetch(
      'https://www.googleapis.com/calendar/v3/calendars/primary/events?sendUpdates=none',
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(eventBody),
      }
    );
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `API error: ${res.status}`);
    }
    return res.json();
  }

  // ============================================================
  // ② ③ Main Search（期間計算修正 + 絶対指定対応）
  // ============================================================
  getSearchRange() {
    const now = new Date();
    let timeMin, timeMax;

    if (this.settings.rangeMode === 'absolute') {
      // ③ 絶対日付指定
      const startStr = this.dateStart.value;
      const endStr = this.dateEnd.value;
      if (!startStr || !endStr) {
        throw new Error('開始日と終了日を指定してください');
      }
      timeMin = new Date(startStr + 'T00:00:00');
      timeMax = new Date(endStr + 'T23:59:59');

      // 過去の開始日は今に補正
      if (timeMin < now) {
        timeMin = new Date(now);
        timeMin.setMinutes(0, 0, 0);
        timeMin.setHours(timeMin.getHours() + 1);
      }
    } else {
      // 相対指定
      timeMin = new Date(now);
      timeMin.setMinutes(0, 0, 0);
      timeMin.setHours(timeMin.getHours() + 1);

      timeMax = new Date(now);
      timeMax.setDate(timeMax.getDate() + this.settings.searchRange);
      timeMax.setHours(23, 59, 59, 0);
    }

    console.log(`[日程調整ツール] 検索範囲: ${timeMin.toLocaleString()} 〜 ${timeMax.toLocaleString()}`);
    return { timeMin, timeMax };
  }

  async searchAvailability() {
    if (this.emails.length === 0) return;

    // 検索前に設定をUIから読み取る（保存ボタン押さなくても反映）
    this.settings.rangeMode = this.rangeMode.value;
    this.settings.searchRange = parseInt(this.searchRange.value);
    this.settings.startTime = this.timeStart.value;
    this.settings.endTime = this.timeEnd.value;
    this.settings.meetingDuration = parseInt(this.meetingDuration.value);

    const rawKeywords = this.excludeKeywords.value;
    this.settings.excludeKeywords = rawKeywords
      .split(/[,、，]/)
      .map((k) => k.trim())
      .filter(Boolean);

    const dayCheckboxes = document.querySelectorAll('.day-check input');
    this.settings.activeDays = [];
    dayCheckboxes.forEach((cb) => {
      if (cb.checked) this.settings.activeDays.push(parseInt(cb.value));
    });

    this.showLoading(true);
    this.results.classList.add('hidden');
    this.conflictsSection.classList.add('hidden');
    this.partialSection.classList.add('hidden');
    this.registerPanel.classList.add('hidden');
    this.registerStatus.classList.add('hidden');

    try {
      const { timeMin, timeMax } = this.getSearchRange();

      const allEvents = await this.getAllEvents(timeMin, timeMax);

      console.log(`[日程調整ツール] 取得イベント数: ${allEvents.length}`);
      for (const ev of allEvents) {
        console.log(`  ${ev.email}: 「${ev.title}」 ${ev.start.toLocaleString()} - ${ev.end.toLocaleString()}`);
      }

      const filteredEvents = allEvents.filter(
        (ev) => !this.shouldExcludeEvent(ev.title)
      );
      console.log(`[日程調整ツール] 除外後イベント数: ${filteredEvents.length}`);

      const busyPeriods = filteredEvents.map((ev) => ({
        email: ev.email,
        start: ev.start,
        end: ev.end,
      }));

      const { freeSlots, partialSlots } = this.findAllSlots(
        timeMin, timeMax, busyPeriods
      );

      this.lastFreeSlots = freeSlots;
      this.lastPartialSlots = partialSlots;

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
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;

    // FreeBusy API をメインに使う
    try {
      const fbData = await this.apiQueryFreeBusy({
        timeMin: timeMin.toISOString(),
        timeMax: timeMax.toISOString(),
        timeZone: tz,
        items: this.emails.map((email) => ({ id: email })),
      });

      for (const email of this.emails) {
        const cal = fbData.calendars?.[email];
        if (cal?.errors) {
          console.warn(`FreeBusy error for ${email}:`, cal.errors);
        }
        if (cal?.busy) {
          for (const b of cal.busy) {
            allEvents.push({
              email,
              title: '',
              start: new Date(b.start),
              end: new Date(b.end),
              source: 'freebusy',
            });
          }
        }
      }
      console.log(`[日程調整ツール] FreeBusy で ${allEvents.length} 件のbusy期間を取得`);
    } catch (e) {
      console.error('FreeBusy API failed:', e.message);
    }

    // Events API で補完（タイトル取得 + 未回答予定の追加）
    for (const email of this.emails) {
      try {
        const data = await this.apiGetEvents(
          email, timeMin.toISOString(), timeMax.toISOString()
        );
        if (data.items) {
          for (const event of data.items) {
            if (event.status === 'cancelled') continue;
            if (event.transparency === 'transparent') continue;

            const attendee = event.attendees?.find(
              (a) => a.email?.toLowerCase() === email.toLowerCase() || a.self
            );
            if (attendee && attendee.responseStatus === 'declined') continue;

            let evStart, evEnd, evTitle;
            if (event.start?.dateTime) {
              evStart = new Date(event.start.dateTime);
              evEnd = new Date(event.end.dateTime);
              evTitle = event.summary || '(タイトルなし)';
            } else if (event.start?.date) {
              const eventDate = new Date(event.start.date);
              const [sh, sm] = this.settings.startTime.split(':').map(Number);
              const [eh, em] = this.settings.endTime.split(':').map(Number);
              evStart = new Date(eventDate);
              evStart.setHours(sh, sm, 0, 0);
              evEnd = new Date(eventDate);
              evEnd.setHours(eh, em, 0, 0);
              evTitle = event.summary || '(終日予定)';
            } else {
              continue;
            }

            let matched = false;
            for (const ev of allEvents) {
              if (ev.email === email && ev.source === 'freebusy') {
                if (evStart < ev.end && evEnd > ev.start) {
                  if (!ev.title) ev.title = evTitle;
                  matched = true;
                  break;
                }
              }
            }

            if (!matched) {
              allEvents.push({
                email,
                title: evTitle,
                start: evStart,
                end: evEnd,
                source: 'events-api',
              });
            }
          }
        }
      } catch (e) {
        console.log(`Events API failed for ${email} (OK - using FreeBusy data): ${e.message}`);
      }
    }

    for (const ev of allEvents) {
      if (!ev.title) {
        ev.title = '(予定あり)';
      }
    }

    return allEvents;
  }

  // ============================================================
  // ② スロット算出（曜日・期間バグ修正）
  // ============================================================
  findAllSlots(timeMin, timeMax, busyPeriods) {
    const freeSlots = [];
    const partialSlots = [];
    const [startHour, startMin] = this.settings.startTime.split(':').map(Number);
    const [endHour, endMin] = this.settings.endTime.split(':').map(Number);
    const duration = this.settings.meetingDuration;
    const slotStep = duration;
    const totalPeople = this.emails.length;
    const now = new Date();

    // ② 開始日を正しく計算: timeMin の日付の検索開始時刻から
    const current = new Date(timeMin);
    // timeMin が今日の場合、今の時間帯の途中かもしれないので
    // まず当日の検索開始時刻を設定
    const dayStart = new Date(current);
    dayStart.setHours(startHour, startMin, 0, 0);

    if (dayStart >= timeMin) {
      // 検索開始時刻がまだ来ていない → その時刻から開始
      current.setHours(startHour, startMin, 0, 0);
    } else {
      // 検索開始時刻は過ぎている → timeMin そのままだが次のスロット区切りに合わせる
      const minOfDay = current.getHours() * 60 + current.getMinutes();
      const startMinOfDay = startHour * 60 + startMin;
      if (minOfDay < startMinOfDay) {
        current.setHours(startHour, startMin, 0, 0);
      } else {
        // 現在時刻以降の次のスロット区切りに合わせる
        const elapsed = minOfDay - startMinOfDay;
        const nextSlotOffset = Math.ceil(elapsed / slotStep) * slotStep;
        current.setHours(startHour, startMin, 0, 0);
        current.setMinutes(current.getMinutes() + nextSlotOffset);
      }
    }

    const endTimeMinutes = endHour * 60 + endMin;
    let safety = 0;
    const maxIterations = 10000;

    while (current < timeMax && safety < maxIterations) {
      safety++;
      const dayOfWeek = current.getDay();

      // ② 曜日フィルタ: activeDays に含まれていない日はスキップ
      if (!this.settings.activeDays.includes(dayOfWeek)) {
        current.setDate(current.getDate() + 1);
        current.setHours(startHour, startMin, 0, 0);
        continue;
      }

      const currentMinOfDay = current.getHours() * 60 + current.getMinutes();

      // 検索時間帯の開始前なら開始時刻にジャンプ
      if (currentMinOfDay < startHour * 60 + startMin) {
        current.setHours(startHour, startMin, 0, 0);
        continue;
      }

      const slotStart = new Date(current);
      const slotEnd = new Date(current);
      slotEnd.setMinutes(slotEnd.getMinutes() + duration);

      const slotEndMinutes = slotEnd.getHours() * 60 + slotEnd.getMinutes();

      // 検索時間帯の終了を超えた or 日をまたいだ → 次の日へ
      if (slotEndMinutes > endTimeMinutes || slotEnd.getDate() !== slotStart.getDate()) {
        current.setDate(current.getDate() + 1);
        current.setHours(startHour, startMin, 0, 0);
        continue;
      }

      // 過去のスロットはスキップ
      if (slotStart > now) {
        const conflicting = busyPeriods.filter(
          (busy) => slotStart < busy.end && slotEnd > busy.start
        );

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
          partialSlots.push({
            start: new Date(slotStart),
            end: new Date(slotEnd),
            conflictCount,
            conflictingEmails: [...conflictingPeople],
            conflictingEvents: conflicting,
          });
        }
      }

      current.setMinutes(current.getMinutes() + slotStep);
    }

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

    if (freeSlots.length === 0) {
      this.freeSlots.innerHTML = `
        <div class="no-results">
          <p>競合なしの空き時間は見つかりませんでした。</p>
          <p style="font-size: 12px; margin-top: 4px;">下の「競合あり」セクションを確認してください。</p>
        </div>`;
    } else {
      this.freeSlots.innerHTML = this.renderSlotCards(freeSlots, 'free');
      this.attachSlotInteractions(this.freeSlots);
    }

    if (partialSlots.length > 0) {
      this.partialSection.classList.remove('hidden');
      this.partialSlots.innerHTML = this.renderSlotCards(partialSlots, 'partial');
      this.attachSlotInteractions(this.partialSlots);
    } else {
      this.partialSection.classList.add('hidden');
    }

    if (freeSlots.length > 0 || partialSlots.length > 0) {
      this.registerPanel.classList.remove('hidden');
      this.updateRegisterButton();
    }

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
    let slotIdx = 0;

    for (const [dateStr, daySlots] of Object.entries(grouped)) {
      html += `<div class="slot-date">${dateStr}</div>`;
      for (const slot of daySlots) {
        const timeStr = this.formatTimeRange(slot.start, slot.end);
        const copyText = `${dateStr} ${timeStr}`;
        const slotId = `${type}-${slotIdx}`;
        const dataStart = slot.start.toISOString();
        const dataEnd = slot.end.toISOString();

        let badge = '';
        if (type === 'partial') {
          const names = slot.conflictingEmails.map((e) => e.split('@')[0]).join(', ');
          badge = `<span class="slot-conflict-badge">${slot.conflictCount}人競合 (${this.escapeHtml(names)})</span>`;
        }

        html += `
          <div class="slot-card slot-card-${type}">
            <div class="slot-header">
              <label class="slot-check-label">
                <input type="checkbox" class="slot-checkbox" data-slot-id="${slotId}" data-start="${dataStart}" data-end="${dataEnd}">
                <div>
                  <span class="slot-time">${timeStr}</span>
                  ${badge}
                </div>
              </label>
              <button class="btn btn-copy" data-copy="${this.escapeHtml(copyText)}">コピー</button>
            </div>
          </div>`;
        slotIdx++;
      }
    }

    return html;
  }

  attachSlotInteractions(container) {
    container.querySelectorAll('.btn-copy').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        navigator.clipboard.writeText(btn.dataset.copy);
        btn.textContent = 'コピー済み';
        setTimeout(() => { btn.textContent = 'コピー'; }, 1500);
      });
    });

    container.querySelectorAll('.slot-checkbox').forEach((cb) => {
      cb.addEventListener('change', () => this.onSlotCheckChanged());
    });
  }

  // ============================================================
  // チェックボックス → 登録パネル
  // ============================================================
  getCheckedSlots() {
    const checked = [];
    document.querySelectorAll('.slot-checkbox:checked').forEach((cb) => {
      checked.push({
        start: cb.dataset.start,
        end: cb.dataset.end,
      });
    });
    return checked;
  }

  onSlotCheckChanged() {
    const checked = this.getCheckedSlots();
    const count = checked.length;

    if (count > 0) {
      const lines = checked.map((s) => {
        const start = new Date(s.start);
        const end = new Date(s.end);
        return `${this.formatDate(start)} ${this.formatTimeRange(start, end)}`;
      });
      this.registerSummary.innerHTML = `
        <strong>${count}件選択中:</strong>
        <ul>${lines.map((l) => `<li>${l}</li>`).join('')}</ul>`;
    } else {
      this.registerSummary.innerHTML = '';
    }

    this.updateRegisterButton();
  }

  updateRegisterButton() {
    const checked = this.getCheckedSlots();
    const title = this.eventTitle.value.trim();
    this.registerBtn.disabled = checked.length === 0 || title.length === 0;
  }

  // ============================================================
  // Google Calendar 予定登録
  // ============================================================
  async registerEvents() {
    const checked = this.getCheckedSlots();
    const title = this.eventTitle.value.trim();
    if (checked.length === 0 || !title) return;

    this.registerBtn.disabled = true;
    this.registerBtn.textContent = '登録中...';
    this.registerStatus.classList.add('hidden');

    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    const attendees = this.emails.map((email) => ({ email }));

    const results = [];

    for (const slot of checked) {
      const eventBody = {
        summary: title,
        start: { dateTime: slot.start, timeZone: tz },
        end: { dateTime: slot.end, timeZone: tz },
        attendees,
      };

      try {
        await this.apiInsertEvent(eventBody);
        const s = new Date(slot.start);
        const e = new Date(slot.end);
        results.push({ success: true, label: `${this.formatDate(s)} ${this.formatTimeRange(s, e)}` });
      } catch (err) {
        const s = new Date(slot.start);
        const e = new Date(slot.end);
        results.push({ success: false, label: `${this.formatDate(s)} ${this.formatTimeRange(s, e)}`, error: err.message });
      }
    }

    const successCount = results.filter((r) => r.success).length;
    const failCount = results.filter((r) => !r.success).length;

    let statusHtml = '';
    if (successCount > 0) {
      statusHtml += `<div class="register-success">✅ ${successCount}件の予定を登録しました！</div>`;
    }
    if (failCount > 0) {
      statusHtml += `<div class="register-error">❌ ${failCount}件の登録に失敗しました:</div>`;
      for (const r of results.filter((r) => !r.success)) {
        statusHtml += `<div class="register-error-detail">${r.label}: ${this.escapeHtml(r.error)}</div>`;
      }
    }

    this.registerStatus.innerHTML = statusHtml;
    this.registerStatus.classList.remove('hidden');

    document.querySelectorAll('.slot-checkbox:checked').forEach((cb) => {
      cb.checked = false;
    });
    this.registerSummary.innerHTML = '';

    this.registerBtn.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="margin-right: 8px;">
        <path d="M12 5v14M5 12h14" stroke-linecap="round"/>
      </svg>
      選択した日程をカレンダーに登録`;
    this.updateRegisterButton();
  }

  // ============================================================
  // Error / Loading
  // ============================================================
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
