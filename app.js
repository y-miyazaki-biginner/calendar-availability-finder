// app.js - Calendar Availability Finder (Web App version)
// Google Identity Services (GIS) を使ったブラウザOAuth2フロー

// =====================================================
// ★ デプロイ時にここを書き換える
// =====================================================
const CONFIG = {
  CLIENT_ID: '416943777269-ie7jg6j4tr53j1lqfplvcnhde0rajuls.apps.googleusercontent.com',
  SCOPES: 'https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.freebusy https://www.googleapis.com/auth/contacts.other.readonly https://www.googleapis.com/auth/directory.readonly',
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
      recurringMode: 'off',    // 'off', 'weekly', 'biweekly'
      recurringWeeks: 4,       // 何週間分チェックするか
    };
    this.lastFreeSlots = [];
    this.lastPartialSlots = [];
    // 追加検索: キャッシュ
    this._cachedBusyPeriods = [];
    this._cachedConflicts = [];
    this._cachedAllSlots = [];
    this._cachedTimeMin = null;
    this._cachedTimeMax = null;
    this._currentMaxOverlap = 0; // 現在表示済みの最大被り件数

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
  // ① ログイン保持: token を localStorage に保存しブラウザを閉じても復元
  // ② トークン期限切れ時に自動でリフレッシュ（再ログイン不要）
  // ============================================================
  initGoogleAuth() {
    // localStorage から token を復元
    const savedToken = localStorage.getItem('calendarToken');
    if (savedToken) {
      this.token = savedToken;
      // token が有効か確認
      this.validateToken().then((valid) => {
        if (valid) {
          this.showMain();
        } else {
          // トークン期限切れ → GIS初期化完了後に自動リフレッシュ
          localStorage.removeItem('calendarToken');
          this.token = null;
          this._needsAutoRefresh = true;
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
            this._peopleApiDisabled = false; // トークン更新でPeople APIを再有効化
            localStorage.setItem('calendarToken', this.token);
            this.showMain();
          },
        });
        // 保存済みトークンが期限切れだった場合、自動でリフレッシュ
        if (this._needsAutoRefresh) {
          this._needsAutoRefresh = false;
          this.refreshToken();
        }
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

  // トークンを自動リフレッシュ（prompt: '' でポップアップなしで試みる）
  refreshToken() {
    if (!this.tokenClient) return;
    try {
      this.tokenClient.requestAccessToken({ prompt: '' });
    } catch (e) {
      console.log('Auto-refresh failed, user needs to re-login:', e.message);
    }
  }

  // API呼び出しのラッパー: 401/403時にトークンをリフレッシュしてリトライ
  async fetchWithAuth(url, options = {}) {
    const headers = { ...options.headers, 'Authorization': `Bearer ${this.token}` };
    let res = await fetch(url, { ...options, headers });

    if (res.status === 401) {
      // トークン期限切れ → リフレッシュして1回だけリトライ
      console.log('Token expired, refreshing...');
      await this.refreshTokenAndWait();
      headers['Authorization'] = `Bearer ${this.token}`;
      res = await fetch(url, { ...options, headers });
    }

    return res;
  }

  // リフレッシュしてトークンが更新されるのを待つ
  refreshTokenAndWait() {
    return new Promise((resolve) => {
      const origCallback = this.tokenClient.callback;
      this.tokenClient.callback = (response) => {
        if (!response.error) {
          this.token = response.access_token;
          this._peopleApiDisabled = false;
          localStorage.setItem('calendarToken', this.token);
        }
        this.tokenClient.callback = origCallback;
        resolve();
      };
      try {
        this.tokenClient.requestAccessToken({ prompt: '' });
      } catch {
        this.tokenClient.callback = origCallback;
        resolve();
      }
    });
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
    this.emailSuggestions = document.getElementById('email-suggestions');
    // 一括操作
    this.bulkActions = document.getElementById('bulk-actions');
    this.selectAllBtn = document.getElementById('select-all-btn');
    this.deselectAllBtn = document.getElementById('deselect-all-btn');
    this.copyCheckedBtn = document.getElementById('copy-checked-btn');
    this.copyCheckedLabel = document.getElementById('copy-checked-label');
    // ユニークアドレスカウンター
    this.uniqueEmailCounter = document.getElementById('unique-email-counter');
    // 定例検索
    this.recurringMode = document.getElementById('recurring-mode');
    this.recurringWeeks = document.getElementById('recurring-weeks');
    this.recurringWeeksGroup = document.getElementById('recurring-weeks-group');
    // 追加検索する
    this.moreSearchSection = document.getElementById('more-search-section');
    this.moreSearchBtn = document.getElementById('more-search-btn');
    this.moreSearchLabel = document.getElementById('more-search-label');
    this.moreSearchHint = document.getElementById('more-search-hint');
    this.additionalResultsSection = document.getElementById('additional-results-section');
  }

  bindEvents() {
    this.authBtn.addEventListener('click', () => this.authenticate());
    this.logoutBtn.addEventListener('click', () => this.logout());
    this.settingsToggle.addEventListener('click', () => this.toggleSettings());
    this.saveSettingsBtn.addEventListener('click', () => this.saveSettings());
    this.addEmailBtn.addEventListener('click', () => this.addEmail());
    this.emailInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        // サジェストが表示中ならハイライト項目を選択
        const highlighted = this.emailSuggestions.querySelector('.suggestion-item.highlighted');
        if (highlighted && !this.emailSuggestions.classList.contains('hidden')) {
          highlighted.click();
          e.preventDefault();
          return;
        }
        this.addEmail();
      } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        this.navigateSuggestions(e.key === 'ArrowDown' ? 1 : -1);
        e.preventDefault();
      } else if (e.key === 'Escape') {
        this.hideSuggestions();
      }
    });
    this.emailInput.addEventListener('input', () => this.onEmailInputChange());
    // サジェスト外クリックで閉じる
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.email-input-wrapper')) {
        this.hideSuggestions();
      }
    });
    this.searchBtn.addEventListener('click', () => this.searchAvailability());
    this.registerBtn.addEventListener('click', () => this.registerEvents());
    this.eventTitle.addEventListener('input', () => this.updateRegisterButton());
    // 範囲モード切替
    this.rangeMode.addEventListener('change', () => this.onRangeModeChange());
    // グループ保存
    this.saveGroupBtn.addEventListener('click', () => this.saveCurrentGroup());
    // 定例検索モード切替
    if (this.recurringMode) {
      this.recurringMode.addEventListener('change', () => this.onRecurringModeChange());
    }
    // 一括操作
    this.selectAllBtn.addEventListener('click', () => this.selectAllSlots());
    this.deselectAllBtn.addEventListener('click', () => this.deselectAllSlots());
    this.copyCheckedBtn.addEventListener('click', () => this.copyCheckedSlots());
    // 追加検索する
    this.moreSearchBtn.addEventListener('click', () => this.expandSearch());
    // セクション内 全選択/全解除（イベント委任）
    this.results.addEventListener('click', (e) => {
      const selectBtn = e.target.closest('.section-select-all');
      const deselectBtn = e.target.closest('.section-deselect-all');
      if (selectBtn) {
        const targetId = selectBtn.dataset.target;
        const container = targetId
          ? document.getElementById(targetId)
          : selectBtn.closest('.additional-results-group')?.querySelector('.slots-list');
        if (container) {
          container.querySelectorAll('.slot-checkbox').forEach((cb) => { cb.checked = true; });
          this.onSlotCheckChange();
        }
      }
      if (deselectBtn) {
        const targetId = deselectBtn.dataset.target;
        const container = targetId
          ? document.getElementById(targetId)
          : deselectBtn.closest('.additional-results-group')?.querySelector('.slots-list');
        if (container) {
          container.querySelectorAll('.slot-checkbox').forEach((cb) => { cb.checked = false; });
          this.onSlotCheckChange();
        }
      }
    });
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
    localStorage.removeItem('calendarToken');
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
    this.updateUniqueEmailCounter();
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

    // 定例検索設定
    if (this.recurringMode) {
      this.settings.recurringMode = this.recurringMode.value;
      this.settings.recurringWeeks = parseInt(this.recurringWeeks.value);
    }

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

    // 定例検索
    if (this.recurringMode) {
      this.recurringMode.value = this.settings.recurringMode || 'off';
      this.recurringWeeks.value = this.settings.recurringWeeks || 4;
      this.onRecurringModeChange();
    }

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

  onRecurringModeChange() {
    if (!this.recurringMode) return;
    const mode = this.recurringMode.value;
    if (mode === 'off') {
      this.recurringWeeksGroup.classList.add('hidden');
    } else {
      this.recurringWeeksGroup.classList.remove('hidden');
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

  // 過去に入力したアドレスの履歴（localStorage保存）
  loadEmailHistory() {
    try {
      const history = JSON.parse(localStorage.getItem('emailHistory') || '[]');
      this.emailHistory = history;
    } catch {
      this.emailHistory = [];
    }
  }

  addToEmailHistory(email) {
    if (!this.emailHistory.includes(email)) {
      this.emailHistory.push(email);
      localStorage.setItem('emailHistory', JSON.stringify(this.emailHistory));
      this.updateUniqueEmailCounter();
    }
  }

  // ============================================================
  // People API サジェスト
  // ============================================================
  onEmailInputChange() {
    const query = this.emailInput.value.trim();
    if (query.length < 2) {
      this.hideSuggestions();
      return;
    }
    // デバウンス
    clearTimeout(this._suggestTimer);
    this._suggestTimer = setTimeout(() => this.fetchSuggestions(query), 300);
  }

  async fetchSuggestions(query) {
    const suggestions = [];
    const lowerQuery = query.toLowerCase();
    const addIfNew = (entry) => {
      if (!this.emails.includes(entry.email) && !suggestions.find((s) => s.email === entry.email)) {
        suggestions.push(entry);
      }
    };

    // 1. 過去に入力したアドレス（emailHistory）から検索
    for (const email of this.emailHistory) {
      if (email.toLowerCase().includes(lowerQuery)) {
        addIfNew({ email, name: '' });
      }
    }

    // 2. 保存済みグループのメンバーから検索
    for (const group of (this.savedGroups || [])) {
      for (const email of group.emails) {
        if (email.toLowerCase().includes(lowerQuery) ||
            group.name.toLowerCase().includes(lowerQuery)) {
          addIfNew({ email, name: `${group.name}` });
        }
      }
    }

    // 3. People APIキャッシュから検索（名前でも検索可）
    for (const entry of (this._contactCache || [])) {
      const match = entry.email.toLowerCase().includes(lowerQuery) ||
                    (entry.name && entry.name.toLowerCase().includes(lowerQuery));
      if (match) addIfNew(entry);
    }

    // 4. People API（Directory）で検索（エラーでも無視）
    if (this.token) {
      try {
        const results = await this.searchPeopleAPI(query);
        for (const r of results) addIfNew(r);
      } catch (e) {
        console.log('People API search failed (non-critical):', e.message);
      }
    }

    this.showSuggestions(suggestions.slice(0, 8));
  }

  async searchPeopleAPI(query) {
    // People API が無効（GCPで有効にしていない）場合は即リターン
    if (this._peopleApiDisabled) return [];

    const results = [];

    // People API - otherContacts (やり取りしたことがある人)
    try {
      const params = new URLSearchParams({
        query,
        readMask: 'names,emailAddresses',
        pageSize: '10',
      });
      const res = await this.fetchWithAuth(
        `https://people.googleapis.com/v1/otherContacts:search?${params}`
      );
      if (res.ok) {
        const data = await res.json();
        for (const r of (data.results || [])) {
          const person = r.person;
          const email = person?.emailAddresses?.[0]?.value;
          const name = person?.names?.[0]?.displayName || '';
          if (email) results.push({ email: email.toLowerCase(), name });
        }
      } else if (res.status === 403) {
        // People API がGCPで有効でない → 以降のリクエストをスキップ
        console.log('People API not enabled in GCP - using local data only');
        this._peopleApiDisabled = true;
        return [];
      }
    } catch { /* network error - ignore */ }

    // People API - directory (組織内ディレクトリ - Google Workspace の場合)
    try {
      const params = new URLSearchParams({
        query,
        readMask: 'names,emailAddresses',
        pageSize: '10',
        sources: 'DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE',
      });
      const res = await this.fetchWithAuth(
        `https://people.googleapis.com/v1/people:searchDirectoryPeople?${params}`
      );
      if (res.ok) {
        const data = await res.json();
        for (const person of (data.people || [])) {
          const email = person?.emailAddresses?.[0]?.value;
          const name = person?.names?.[0]?.displayName || '';
          if (email && !results.find((r) => r.email === email.toLowerCase())) {
            results.push({ email: email.toLowerCase(), name });
          }
        }
      }
      // directory API の 403 は無視（Workspace以外では使えない）
    } catch { /* network error - ignore */ }

    // コンタクトキャッシュに追加
    if (!this._contactCache) this._contactCache = [];
    for (const r of results) {
      if (!this._contactCache.find((c) => c.email === r.email)) {
        this._contactCache.push(r);
      }
    }

    return results;
  }

  showSuggestions(suggestions) {
    if (suggestions.length === 0) {
      this.hideSuggestions();
      return;
    }

    this.emailSuggestions.innerHTML = suggestions
      .map((s, i) => `
        <div class="suggestion-item" data-email="${this.escapeHtml(s.email)}" data-index="${i}">
          <div class="suggestion-name">${s.name ? this.escapeHtml(s.name) : '<span class="suggestion-no-name">名前なし</span>'}</div>
          <div class="suggestion-email">${this.escapeHtml(s.email)}</div>
        </div>`)
      .join('');

    this.emailSuggestions.classList.remove('hidden');

    this.emailSuggestions.querySelectorAll('.suggestion-item').forEach((item) => {
      item.addEventListener('click', () => {
        this.emailInput.value = item.dataset.email;
        this.hideSuggestions();
        this.addEmail();
      });
      item.addEventListener('mouseenter', () => {
        this.emailSuggestions.querySelectorAll('.suggestion-item').forEach((el) => el.classList.remove('highlighted'));
        item.classList.add('highlighted');
      });
    });
  }

  hideSuggestions() {
    this.emailSuggestions.classList.add('hidden');
    this.emailSuggestions.innerHTML = '';
  }

  navigateSuggestions(direction) {
    const items = this.emailSuggestions.querySelectorAll('.suggestion-item');
    if (items.length === 0) return;

    const current = this.emailSuggestions.querySelector('.suggestion-item.highlighted');
    let idx = -1;
    if (current) {
      idx = parseInt(current.dataset.index);
      current.classList.remove('highlighted');
    }

    idx += direction;
    if (idx < 0) idx = items.length - 1;
    if (idx >= items.length) idx = 0;

    items[idx].classList.add('highlighted');
    items[idx].scrollIntoView({ block: 'nearest' });
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
    this.updateUniqueEmailCounter();
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
    this.updateUniqueEmailCounter();
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
  // ユニークアドレス数カウンター（参考情報）
  // ============================================================
  updateUniqueEmailCounter() {
    const allEmails = new Set();
    // 過去に入力したアドレス
    for (const email of this.emailHistory) {
      allEmails.add(email.toLowerCase());
    }
    // 保存済みグループのメンバー
    for (const group of (this.savedGroups || [])) {
      for (const email of group.emails) {
        allEmails.add(email.toLowerCase());
      }
    }
    const count = allEmails.size;
    if (this.uniqueEmailCounter) {
      this.uniqueEmailCounter.textContent = `このツールで使用したユニークアドレス数: ${count}`;
    }
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
  // イベントの時間帯が検索設定の時間帯と重なっているか確認
  // ============================================================
  isWithinSearchTimeRange(eventStart, eventEnd) {
    const [startH, startM] = this.settings.startTime.split(':').map(Number);
    const [endH, endM] = this.settings.endTime.split(':').map(Number);
    const searchStartMin = startH * 60 + startM;
    const searchEndMin = endH * 60 + endM;

    // イベントが日をまたぐ場合（終日予定をFreeBusyが返した場合など）は
    // 検索時間帯と必ず重なるのでtrue
    const durationMs = eventEnd.getTime() - eventStart.getTime();
    const durationHours = durationMs / (1000 * 60 * 60);
    if (durationHours >= 24) {
      return true;
    }

    const eventStartMin = eventStart.getHours() * 60 + eventStart.getMinutes();
    let eventEndMin = eventEnd.getHours() * 60 + eventEnd.getMinutes();

    // 日をまたぐイベント（例: 23:00〜翌01:00）や
    // 終日予定の端数（00:00終了）を処理
    if (eventEndMin === 0 && eventEnd.getDate() !== eventStart.getDate()) {
      eventEndMin = 24 * 60; // 24:00として扱う
    }

    // イベントが検索時間帯と重なっているか
    return eventEndMin > searchStartMin && eventStartMin < searchEndMin;
  }

  // ============================================================
  // Google Calendar API
  // ============================================================
  async apiQueryFreeBusy(params) {
    const res = await this.fetchWithAuth('https://www.googleapis.com/calendar/v3/freeBusy', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
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
    const res = await this.fetchWithAuth(
      `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(calendarId)}/events?${params}`
    );
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `API error: ${res.status}`);
    }
    return res.json();
  }

  async apiInsertEvent(eventBody) {
    const res = await this.fetchWithAuth(
      'https://www.googleapis.com/calendar/v3/calendars/primary/events?sendUpdates=none',
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
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

    // 定例検索モードの場合、チェック週数分の範囲を確保
    const recurringMode = this.settings.recurringMode || 'off';
    if (recurringMode !== 'off') {
      const recurringWeeks = this.settings.recurringWeeks || 4;
      const weekInterval = recurringMode === 'biweekly' ? 2 : 1;
      const requiredDays = recurringWeeks * weekInterval * 7;
      const requiredMax = new Date(timeMin);
      requiredMax.setDate(requiredMax.getDate() + requiredDays);
      requiredMax.setHours(23, 59, 59, 0);
      if (requiredMax > timeMax) {
        timeMax = requiredMax;
        console.log(`[定例検索] 検索範囲を${recurringWeeks * weekInterval}週間分に拡張`);
      }
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

    // 定例検索設定
    if (this.recurringMode) {
      this.settings.recurringMode = this.recurringMode.value;
      this.settings.recurringWeeks = parseInt(this.recurringWeeks.value);
    }

    this.showLoading(true);
    this.results.classList.add('hidden');
    this.bulkActions.classList.add('hidden');
    this.conflictsSection.classList.add('hidden');
    this.partialSection.classList.add('hidden');
    this.registerPanel.classList.add('hidden');
    this.registerStatus.classList.add('hidden');
    this.moreSearchSection.classList.add('hidden');
    this.additionalResultsSection.classList.add('hidden');
    this.additionalResultsSection.innerHTML = '';

    try {
      const { timeMin, timeMax } = this.getSearchRange();

      const allEvents = await this.getAllEvents(timeMin, timeMax);

      console.log(`\n[日程調整ツール] 合計取得イベント数: ${allEvents.length} (FreeBusy + Events API)`);

      const filteredEvents = allEvents.filter(
        (ev) => !this.shouldExcludeEvent(ev.title)
      );

      const excludedEvents = allEvents.filter((ev) => this.shouldExcludeEvent(ev.title));
      if (excludedEvents.length > 0) {
        console.log(`\n[除外キーワード] ${excludedEvents.length}件を除外:`);
        for (const ev of excludedEvents) {
          console.log(`  ❌ ${ev.email}: 「${ev.title}」 ${ev.start.toLocaleString()} - ${ev.end.toLocaleString()}`);
        }
      }
      console.log(`\n[日程調整ツール] 最終busy期間: ${filteredEvents.length}件`);
      for (const ev of filteredEvents) {
        console.log(`  ✅ ${ev.email}: 「${ev.title}」 ${ev.start.toLocaleString()} - ${ev.end.toLocaleString()} [${ev.source}]`);
      }

      const busyPeriods = filteredEvents.map((ev) => ({
        email: ev.email,
        start: ev.start,
        end: ev.end,
      }));

      const conflictsInRange = filteredEvents.filter((ev) =>
        ev.start < timeMax && ev.end > timeMin &&          // 検索日付範囲内
        this.settings.activeDays.includes(ev.start.getDay()) && // 対象曜日
        this.isWithinSearchTimeRange(ev.start, ev.end)     // 検索時間帯内
      );

      // キャッシュ保存（「追加検索する」で再利用）
      this._cachedBusyPeriods = busyPeriods;
      this._cachedConflicts = conflictsInRange;
      this._cachedTimeMin = timeMin;
      this._cachedTimeMax = timeMax;
      this._currentMaxOverlap = 0; // 初期閾値（初回検索結果の後に設定される）
      this.additionalResultsSection.innerHTML = '';
      this.additionalResultsSection.classList.add('hidden');

      // 定例検索モードの場合、毎週/隔週で空いている曜日×時間帯を抽出
      const recurringMode = this.settings.recurringMode || 'off';
      if (recurringMode !== 'off') {
        const { freeSlots, partialSlots } = this.findAllSlots(
          timeMin, timeMax, busyPeriods
        );
        const recurringResults = this.findRecurringSlots(freeSlots, partialSlots, timeMin, timeMax);
        this.lastFreeSlots = [];
        this.lastPartialSlots = [];
        this.renderRecurringResults(recurringResults, conflictsInRange);
        this.moreSearchSection.classList.add('hidden');
      } else {
        this.renderInitialResults(busyPeriods, conflictsInRange);
      }
    } catch (error) {
      this.showError(error.message);
    } finally {
      this.showLoading(false);
    }
  }

  async getAllEvents(timeMin, timeMax) {
    const allEvents = [];
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;

    console.log(`\n========================================`);
    console.log(`[日程調整ツール] デバッグ: 検索開始`);
    console.log(`  検索範囲: ${timeMin.toLocaleString()} 〜 ${timeMax.toLocaleString()}`);
    console.log(`  検索時間帯: ${this.settings.startTime} 〜 ${this.settings.endTime}`);
    console.log(`  対象メール: ${this.emails.join(', ')}`);
    console.log(`  除外キーワード: ${(this.settings.excludeKeywords || []).join(', ') || '(なし)'}`);
    console.log(`========================================\n`);

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
          console.warn(`[FreeBusy] ⚠ ${email} エラー:`, cal.errors);
        }
        if (cal?.busy) {
          console.log(`[FreeBusy] ${email}: ${cal.busy.length}件のbusy期間`);
          for (const b of cal.busy) {
            const start = new Date(b.start);
            const end = new Date(b.end);
            const durationHours = (end.getTime() - start.getTime()) / (1000 * 60 * 60);

            console.log(`  📅 ${start.toLocaleString()} 〜 ${end.toLocaleString()} (${durationHours.toFixed(1)}h)`);
            allEvents.push({
              email,
              title: '',
              start,
              end,
              source: 'freebusy',
            });
          }
        } else {
          console.log(`[FreeBusy] ${email}: busy期間なし`);
        }
      }
      console.log(`[日程調整ツール] FreeBusy で ${allEvents.length} 件のbusy期間を取得`);
    } catch (e) {
      console.error('FreeBusy API failed:', e.message);
    }

    // Events API で補完（タイトル取得 + FreeBusy に含まれない予定の追加）
    // ※ 「いいえ（declined）」のみ除外。未回答（needsAction）・仮承諾（tentative）・
    //    attendeesリストにいない場合もすべてbusy扱いにする
    // また、終日予定（start.date）は対象外とし、FreeBusyのbusy期間からも除去する
    //
    // 重要: FreeBusy APIは隣接する予定を連結して1つのbusy期間として返す。
    // Events APIで個別の予定が取得できた場合、FreeBusyのbusy期間を個別イベントで置き換え、
    // 正確なイベント単位での表示を実現する。
    const allDayPeriods = []; // 終日予定のbusy期間（後でFreeBusyから除去用）
    const eventsApiSuccessEmails = new Set(); // Events APIでアクセスできたメール

    for (const email of this.emails) {
      try {
        const data = await this.apiGetEvents(
          email, timeMin.toISOString(), timeMax.toISOString()
        );
        eventsApiSuccessEmails.add(email);
        console.log(`\n[Events API] ${email}: ${data.items?.length || 0}件のイベント取得`);
        if (data.items) {
          for (const event of data.items) {
            if (event.status === 'cancelled') {
              console.log(`  [スキップ] 「${event.summary}」- cancelled`);
              continue;
            }
            if (event.transparency === 'transparent') {
              console.log(`  [スキップ] 「${event.summary}」- transparent (公開設定: 予定あり→外)`);
              continue;
            }

            // 出欠確認: 「いいえ（declined）」の場合のみスキップ
            const attendee = event.attendees?.find(
              (a) => a.email?.toLowerCase() === email.toLowerCase() || a.self
            );
            const responseStatus = attendee ? attendee.responseStatus : 'no-attendee-entry';

            // declined のみ除外。attendee が見つからない場合も busy 扱い
            if (attendee && attendee.responseStatus === 'declined') {
              console.log(`  [除外] ${email}: 「${event.summary}」- declined（いいえ）`);
              continue;
            }

            // 終日予定 (start.date) は対象外にする（祝日・休暇ラベル等）
            // FreeBusyに含まれている可能性があるので、後で除去するためperiodを記録
            if (event.start?.date) {
              console.log(`  [終日除外] 「${event.summary}」- 終日予定のため対象外 (${event.start.date}〜${event.end.date})`);
              // 終日予定のUTC期間を記録（FreeBusyが返す形式と合わせる）
              const adStart = new Date(event.start.date + 'T00:00:00Z');
              const adEnd = new Date(event.end.date + 'T00:00:00Z');
              allDayPeriods.push({ email, start: adStart, end: adEnd });
              continue;
            }

            if (!event.start?.dateTime) continue;

            const evTitle = event.summary || '(タイトルなし)';
            const evStart = new Date(event.start.dateTime);
            const evEnd = new Date(event.end.dateTime);

            console.log(`  [処理] 「${evTitle}」 status=${responseStatus} start=${event.start.dateTime} end=${event.end.dateTime}`);

            // Events APIで取得した個別イベントを直接追加
            // （FreeBusy の連結期間は後で置き換えるので、ここでは重複チェック不要）
            allEvents.push({
              email,
              title: evTitle,
              start: evStart,
              end: evEnd,
              source: 'events-api',
            });
          }
        }
      } catch (e) {
        // Events APIアクセス不可の場合、終日予定の除外もできないが
        // FreeBusyデータをそのまま使う（終日予定がbusyに含まれる可能性あり）
        console.log(`[Events API] ⚠ ${email}: アクセス不可 - FreeBusyデータのみ使用 (${e.message})`);
      }
    }

    // Events APIで個別イベントが取得できた人のFreeBusy連結期間を除去
    // （個別イベントが正確なので、FreeBusyの連結期間は不要）
    // Events APIでアクセスできなかった人のFreeBusy期間はそのまま残す
    {
      const beforeCount = allEvents.length;
      for (let i = allEvents.length - 1; i >= 0; i--) {
        const ev = allEvents[i];
        if (ev.source !== 'freebusy') continue;

        if (eventsApiSuccessEmails.has(ev.email)) {
          // Events APIで取得できた人 → FreeBusy期間を除去（個別イベントで置換済み）
          console.log(`  [FreeBusy→Events置換] ${ev.email} FreeBusy期間を除去: ${ev.start.toLocaleString()} 〜 ${ev.end.toLocaleString()}`);
          allEvents.splice(i, 1);
          continue;
        }

        // Events APIでアクセスできなかった人 →
        // 1. 他人のEvents APIから取得した終日予定(allDayPeriods)に包含されるか
        // 2. UTC midnight境界のbusy期間（= FreeBusy APIが返す終日予定の特徴）か
        let removed = false;

        // 1. 他ユーザーのEvents APIで判明した終日予定の期間に包含される場合
        for (const adp of allDayPeriods) {
          // 終日予定は誰のカレンダーでも共通（祝日等）なので、email問わずチェック
          if (ev.start >= adp.start && ev.end <= adp.end) {
            console.log(`  [終日除去] FreeBusy busy期間を除去（他ユーザーの終日予定と一致）: ${ev.email} ${ev.start.toLocaleString()} 〜 ${ev.end.toLocaleString()}`);
            allEvents.splice(i, 1);
            removed = true;
            break;
          }
        }
        if (removed) continue;

        // 2. UTC midnight境界チェック: FreeBusyが返す終日予定は
        //    UTC 00:00:00 開始かつ UTC 00:00:00 終了で、ちょうど日数分の期間
        //    （例: 1日の終日予定 = 24h, 2日の終日予定 = 48h）
        const startUTC = ev.start;
        const endUTC = ev.end;
        const isStartMidnightUTC = startUTC.getUTCHours() === 0 && startUTC.getUTCMinutes() === 0 && startUTC.getUTCSeconds() === 0;
        const isEndMidnightUTC = endUTC.getUTCHours() === 0 && endUTC.getUTCMinutes() === 0 && endUTC.getUTCSeconds() === 0;
        const durationMs = endUTC.getTime() - startUTC.getTime();
        const isExactDays = durationMs > 0 && durationMs % (24 * 60 * 60 * 1000) === 0;

        if (isStartMidnightUTC && isEndMidnightUTC && isExactDays) {
          console.log(`  [終日除去] FreeBusy busy期間を終日予定として除去（UTC midnight境界）: ${ev.email} ${ev.start.toLocaleString()} 〜 ${ev.end.toLocaleString()}`);
          allEvents.splice(i, 1);
        }
      }
      const removedCount = beforeCount - allEvents.length;
      if (removedCount > 0) {
        console.log(`[FreeBusy整理] ${removedCount}件のFreeBusy期間を除去/置換`);
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
  // 初回検索結果の表示
  // free=被り0、partial=被り予定あり（各人最大1件まで）
  // 「追加検索する」で被り件数の閾値を上げて段階的に候補を追加
  // ============================================================
  renderInitialResults(busyPeriods, conflictsInRange) {
    // 全スロットを被り件数付きで計算してキャッシュ
    this._cachedAllSlots = this.findAllSlotsWithOverlapCount(
      this._cachedTimeMin, this._cachedTimeMax, busyPeriods
    );

    // 初回の閾値: 各参加者につき最大1件の被りまで許容
    // （被り件数 ≤ 被り人数 = 各人最大1件）
    const initialThreshold = this.emails.length;
    this._currentMaxOverlap = initialThreshold;

    // 閾値以下のスロットを free / partial に振り分け
    const freeSlots = [];
    const partialSlots = [];
    for (const slot of this._cachedAllSlots) {
      if (slot.overlapCount > initialThreshold) continue;
      if (slot.conflictCount === 0) {
        freeSlots.push(slot);
      } else {
        partialSlots.push(slot);
      }
    }

    partialSlots.sort((a, b) => {
      if (a.conflictCount !== b.conflictCount) return a.conflictCount - b.conflictCount;
      return a.start - b.start;
    });

    this.lastFreeSlots = freeSlots;
    this.lastPartialSlots = partialSlots;
    this.renderResults(freeSlots, partialSlots, conflictsInRange);

    // 「追加検索する」ボタンの表示/非表示
    this.updateMoreSearchButton();
  }

  // ============================================================
  // 追加検索する（被り件数の閾値を+1して追加候補を表示）
  // ============================================================
  expandSearch() {
    const nextThreshold = this._currentMaxOverlap + 1;

    // 次の閾値で新たに該当するスロットを取得
    const additionalSlots = this._cachedAllSlots.filter(
      (slot) => slot.overlapCount === nextThreshold
    );

    this._currentMaxOverlap = nextThreshold;

    // 追加結果をセクションとして追加
    if (additionalSlots.length > 0) {
      this.additionalResultsSection.classList.remove('hidden');
      this.renderAdditionalSection(additionalSlots, nextThreshold);
    }

    // ボタン更新
    this.updateMoreSearchButton();
  }

  // 「追加検索する」ボタンの表示/非表示制御
  updateMoreSearchButton() {
    // キャッシュ内に現在の閾値を超えるスロットがあるかチェック
    const hasMore = this._cachedAllSlots &&
      this._cachedAllSlots.some((slot) => slot.overlapCount > this._currentMaxOverlap);

    if (hasMore) {
      this.moreSearchSection.classList.remove('hidden');
      this.moreSearchLabel.textContent = '追加検索する';
      const initialThreshold = this.emails.length;
      if (this._currentMaxOverlap > initialThreshold) {
        this.moreSearchHint.textContent = `現在: 被り${this._currentMaxOverlap}件までの候補を表示中`;
      } else {
        this.moreSearchHint.textContent = '被りが多い候補も表示します';
      }
    } else {
      this.moreSearchSection.classList.add('hidden');
    }
  }

  // 全スロットを被り件数付きで計算
  findAllSlotsWithOverlapCount(timeMin, timeMax, busyPeriods) {
    const allSlots = [];
    const [startHour, startMin] = this.settings.startTime.split(':').map(Number);
    const [endHour, endMin] = this.settings.endTime.split(':').map(Number);
    const duration = this.settings.meetingDuration;
    const slotStep = duration;
    const now = new Date();

    const current = new Date(timeMin);
    const dayStart = new Date(current);
    dayStart.setHours(startHour, startMin, 0, 0);

    if (dayStart >= timeMin) {
      current.setHours(startHour, startMin, 0, 0);
    } else {
      const minOfDay = current.getHours() * 60 + current.getMinutes();
      const startMinOfDay = startHour * 60 + startMin;
      if (minOfDay < startMinOfDay) {
        current.setHours(startHour, startMin, 0, 0);
      } else {
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

      if (!this.settings.activeDays.includes(dayOfWeek)) {
        current.setDate(current.getDate() + 1);
        current.setHours(startHour, startMin, 0, 0);
        continue;
      }

      const currentMinOfDay = current.getHours() * 60 + current.getMinutes();
      if (currentMinOfDay < startHour * 60 + startMin) {
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

      if (slotStart > now) {
        const conflicting = busyPeriods.filter(
          (busy) => slotStart < busy.end && slotEnd > busy.start
        );
        const conflictingPeople = new Set(conflicting.map((c) => c.email));
        const conflictCount = conflictingPeople.size;
        const overlapCount = conflicting.length; // 被り予定の合計件数

        allSlots.push({
          start: new Date(slotStart),
          end: new Date(slotEnd),
          conflictCount,         // 被り人数
          overlapCount,          // 被り予定件数
          conflictingEmails: [...conflictingPeople],
          conflictingEvents: conflicting,
        });
      }

      current.setMinutes(current.getMinutes() + slotStep);
    }

    return allSlots;
  }

  // 追加セクションをレンダリング
  renderAdditionalSection(slots, overlapCount) {
    const sectionDiv = document.createElement('div');
    sectionDiv.className = 'additional-results-group';

    // 見出し行（見出し＋全選択/全解除）
    const headingRow = document.createElement('div');
    headingRow.className = 'section-heading-row';

    const heading = document.createElement('h3');
    heading.className = 'results-heading results-heading-additional';
    heading.textContent = `被り予定${overlapCount}件の候補`;
    headingRow.appendChild(heading);

    const btnsDiv = document.createElement('div');
    btnsDiv.className = 'section-select-btns';
    btnsDiv.innerHTML = `
      <button class="btn btn-secondary btn-mini section-select-all">全選択</button>
      <button class="btn btn-secondary btn-mini section-deselect-all">全解除</button>
    `;
    headingRow.appendChild(btnsDiv);
    sectionDiv.appendChild(headingRow);

    const slotsContainer = document.createElement('div');
    slotsContainer.className = 'slots-list';
    slotsContainer.innerHTML = this.renderSlotCards(slots, 'additional');
    sectionDiv.appendChild(slotsContainer);

    this.additionalResultsSection.appendChild(sectionDiv);
    this.attachSlotInteractions(slotsContainer);

    // 一括操作バーが非表示なら表示する
    this.bulkActions.classList.remove('hidden');
    // 登録パネルも表示
    this.registerPanel.classList.remove('hidden');
    this.updateRegisterButton();
  }

  // ============================================================
  // ② スロット算出（曜日・期間バグ修正）
  // ============================================================
  findAllSlots(timeMin, timeMax, busyPeriods, allowedConflicts = 0) {
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

        if (conflictCount <= allowedConflicts) {
          freeSlots.push({
            start: new Date(slotStart),
            end: new Date(slotEnd),
            conflictCount,
            conflictingEmails: conflictCount > 0 ? [...conflictingPeople] : [],
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
  // 定例候補日程の抽出（毎週 / 隔週）
  // ============================================================
  findRecurringSlots(freeSlots, partialSlots, timeMin, timeMax) {
    const mode = this.settings.recurringMode; // 'weekly' or 'biweekly'
    const checkWeeks = this.settings.recurringWeeks || 4;
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];

    // 各曜日が検索範囲内に何回出現するかを正確に計算
    const dayOccurrences = {};  // { 0:日曜の回数, 1:月曜の回数, ... }
    for (let d = 0; d < 7; d++) dayOccurrences[d] = 0;
    const countDay = new Date(timeMin);
    countDay.setHours(0, 0, 0, 0);
    const countEnd = new Date(timeMax);
    countEnd.setHours(23, 59, 59, 0);
    while (countDay <= countEnd) {
      dayOccurrences[countDay.getDay()]++;
      countDay.setDate(countDay.getDate() + 1);
    }

    console.log(`[定例検索] 各曜日の出現回数:`, Object.entries(dayOccurrences).map(([d, c]) => `${dayNames[d]}=${c}`).join(', '));

    // 全スロット（free のみ使用。partialは「全員空き」ではないので定例候補には不適格）
    // ただし partial も「ほぼ空き」として別枠で集計
    const freeByKey = {};   // "dayOfWeek-HH:MM" → 週ごとの出現
    const partialByKey = {};

    const addToMap = (map, slot) => {
      const dow = slot.start.getDay();
      const timeKey = `${String(slot.start.getHours()).padStart(2, '0')}:${String(slot.start.getMinutes()).padStart(2, '0')}`;
      const key = `${dow}-${timeKey}`;
      if (!map[key]) {
        map[key] = {
          dayOfWeek: dow,
          dayName: dayNames[dow],
          timeStart: timeKey,
          timeEnd: `${String(slot.end.getHours()).padStart(2, '0')}:${String(slot.end.getMinutes()).padStart(2, '0')}`,
          dates: [],  // 出現した日付リスト
        };
      }
      // 同じ日付の重複を避ける
      const dateStr = slot.start.toDateString();
      if (!map[key].dates.find((d) => d.dateStr === dateStr)) {
        map[key].dates.push({ dateStr, slot });
      }
    };

    for (const s of freeSlots) addToMap(freeByKey, s);
    // partial も free 側にマージ（「空きまたは一部競合」として集計するため）
    for (const s of partialSlots) addToMap(partialByKey, s);

    // 結果
    const recurringFree = [];   // 全週空き（全員空き）
    const recurringPartial = []; // ほぼ空き

    // freeByKey に存在するキーをベースに判定
    const allKeys = new Set([...Object.keys(freeByKey), ...Object.keys(partialByKey)]);

    for (const key of allKeys) {
      const freeData = freeByKey[key];
      const partialData = partialByKey[key];

      // どちらかのデータから基本情報を取得
      const baseData = freeData || partialData;
      const dow = baseData.dayOfWeek;
      const totalOccurrences = dayOccurrences[dow]; // この曜日が範囲内に出現する回数
      const targetWeeks = Math.min(checkWeeks, totalOccurrences);

      if (targetWeeks <= 0) continue;

      // free（全員空き）で出現した回数
      const freeCount = freeData ? freeData.dates.length : 0;

      if (mode === 'biweekly') {
        // 隔週: 日付を出現順でインデックス付けし、偶数番目/奇数番目でチェック
        // （例: 月曜が3/10,3/17,3/24,3/31の順なら index 0,1,2,3）

        // この曜日のfreeスロットの出現日を日付順に取得
        const freeDatesSorted = freeData
          ? [...freeData.dates].sort((a, b) => a.slot.start - b.slot.start)
          : [];

        // この曜日の全出現日リスト（free/partial/busyすべて含む）を日付順に構築
        // → dayOccurrences[dow] 回分のインデックスを割り当て
        // free出現日のインデックスを特定する
        const allDayDates = [];
        const tmpDay = new Date(timeMin);
        tmpDay.setHours(0, 0, 0, 0);
        while (tmpDay <= countEnd) {
          if (tmpDay.getDay() === dow) {
            allDayDates.push(tmpDay.toDateString());
          }
          tmpDay.setDate(tmpDay.getDate() + 1);
        }

        // freeスロットの日付がallDayDatesの何番目かを調べる
        const freeIndices = new Set();
        for (const fd of freeDatesSorted) {
          const idx = allDayDates.indexOf(fd.dateStr);
          if (idx >= 0) freeIndices.add(idx);
        }

        for (const startOffset of [0, 1]) {
          // 隔週パターン: startOffset, startOffset+2, startOffset+4, ...
          const patternIndices = [];
          for (let i = startOffset; i < allDayDates.length; i += 2) {
            patternIndices.push(i);
          }
          const biweeklyTarget = Math.min(checkWeeks, patternIndices.length);
          if (biweeklyTarget <= 0) continue;

          // このパターンに含まれる週のうち free の数
          const freeInPattern = patternIndices.filter((i) => freeIndices.has(i)).length;

          if (freeInPattern >= biweeklyTarget) {
            recurringFree.push({
              ...baseData,
              matchCount: freeInPattern,
              totalExpected: biweeklyTarget,
              label: startOffset === 0 ? '偶数週' : '奇数週',
            });
          } else if (freeInPattern >= biweeklyTarget - 1) {
            recurringPartial.push({
              ...baseData,
              matchCount: freeInPattern,
              totalExpected: biweeklyTarget,
              label: startOffset === 0 ? '偶数週' : '奇数週',
            });
          }
        }
      } else {
        // 毎週: この曜日×時間帯が targetWeeks 回以上 free か
        if (freeCount >= targetWeeks) {
          recurringFree.push({
            ...baseData,
            matchCount: freeCount,
            totalExpected: targetWeeks,
          });
        } else if (freeCount >= targetWeeks - 1) {
          // 1週だけ不足 → ほぼ空き
          recurringPartial.push({
            ...baseData,
            matchCount: freeCount,
            totalExpected: targetWeeks,
          });
        }
      }
    }

    // 重複排除（隔週で偶数/奇数の両方にマッチした場合）
    const dedup = (arr) => {
      const seen = new Set();
      return arr.filter((item) => {
        const key = `${item.dayOfWeek}-${item.timeStart}-${item.label || ''}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    };

    // 曜日 → 時間帯でソート
    const sortFn = (a, b) => {
      if (a.dayOfWeek !== b.dayOfWeek) return a.dayOfWeek - b.dayOfWeek;
      return a.timeStart.localeCompare(b.timeStart);
    };

    const resultFree = dedup(recurringFree).sort(sortFn);
    const resultPartial = dedup(recurringPartial).sort(sortFn);

    console.log(`[定例検索] mode=${mode} checkWeeks=${checkWeeks}`);
    console.log(`  全週空き=${resultFree.length}件 ほぼ空き=${resultPartial.length}件`);
    for (const r of resultFree) {
      console.log(`  ✅ ${r.dayName}曜 ${r.timeStart}-${r.timeEnd} (${r.matchCount}/${r.totalExpected}週 free) ${r.label || ''}`);
    }
    for (const r of resultPartial) {
      console.log(`  ⚠ ${r.dayName}曜 ${r.timeStart}-${r.timeEnd} (${r.matchCount}/${r.totalExpected}週 free) ${r.label || ''}`);
    }

    return { recurringFree: resultFree, recurringPartial: resultPartial };
  }

  renderRecurringResults(recurringResults, conflictsInRange) {
    const { recurringFree, recurringPartial } = recurringResults;
    const modeLabel = this.settings.recurringMode === 'biweekly' ? '隔週' : '毎週';

    this.results.classList.remove('hidden');
    this.bulkActions.classList.add('hidden'); // 定例モードでは一括操作不要
    this.registerPanel.classList.add('hidden');

    // 競合なしセクション → 定例候補
    if (recurringFree.length === 0) {
      this.freeSlots.innerHTML = `
        <div class="no-results">
          <p>${modeLabel}で全員が空いている時間帯は見つかりませんでした。</p>
          <p style="font-size: 12px; margin-top: 4px;">下の「ほぼ毎週空き」セクションを確認してください。</p>
        </div>`;
    } else {
      this.freeSlots.innerHTML = this.renderRecurringCards(recurringFree, 'free', modeLabel);
    }

    // Section heading update
    const freeHeading = this.freeSection?.querySelector('.results-heading-free');
    if (freeHeading) freeHeading.textContent = `${modeLabel}空きの候補`;

    // 競合少セクション → ほぼ空き
    if (recurringPartial.length > 0) {
      this.partialSection.classList.remove('hidden');
      this.partialSlots.innerHTML = this.renderRecurringCards(recurringPartial, 'partial', modeLabel);
      const partialHeading = this.partialSection?.querySelector('.results-heading-partial');
      if (partialHeading) partialHeading.textContent = `${modeLabel}ほぼ空きの候補（一部週に被り予定あり）`;
    } else {
      this.partialSection.classList.add('hidden');
    }

    // 競合する予定
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

  renderRecurringCards(items, type, modeLabel) {
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
    let html = '';

    // 曜日でグループ化
    const byDay = {};
    for (const item of items) {
      if (!byDay[item.dayOfWeek]) byDay[item.dayOfWeek] = [];
      byDay[item.dayOfWeek].push(item);
    }

    for (const [dow, dayItems] of Object.entries(byDay)) {
      html += `<div class="slot-date">${modeLabel} ${dayNames[parseInt(dow)]}曜日</div>`;
      for (const item of dayItems) {
        const biweeklyLabel = item.label ? ` (${item.label})` : '';
        const badge = type === 'partial'
          ? `<span class="slot-conflict-badge">${item.matchCount}/${item.totalExpected}週空き</span>`
          : `<span class="slot-recurring-badge">${item.matchCount}週連続空き</span>`;

        const copyText = `${modeLabel} ${dayNames[parseInt(dow)]}曜 ${item.timeStart} - ${item.timeEnd}${biweeklyLabel}`;

        html += `
          <div class="slot-card slot-card-${type}">
            <div class="slot-header">
              <div>
                <span class="slot-time">${item.timeStart} - ${item.timeEnd}${biweeklyLabel}</span>
                ${badge}
              </div>
              <button class="btn btn-copy" data-copy="${this.escapeHtml(copyText)}">コピー</button>
            </div>
          </div>`;
      }
    }

    // コピーボタンのイベント付与（renderの後に呼ぶ必要あり）
    setTimeout(() => {
      this.results.querySelectorAll('.btn-copy').forEach((btn) => {
        btn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          navigator.clipboard.writeText(btn.dataset.copy);
          btn.textContent = 'コピー済み';
          setTimeout(() => { btn.textContent = 'コピー'; }, 1500);
        });
      });
    }, 0);

    return html;
  }

  // ============================================================
  // Rendering
  // ============================================================
  renderResults(freeSlots, partialSlots, conflictsInRange) {
    this.results.classList.remove('hidden');

    // 通常モード: 見出しを設定
    const freeHeading = this.freeSection?.querySelector('.results-heading-free');
    if (freeHeading) {
      freeHeading.textContent = '被りなしの候補';
    }
    const partialHeading = this.partialSection?.querySelector('.results-heading-partial');
    if (partialHeading) partialHeading.textContent = '被り予定ありの候補（被りが少ない順）';

    // 一括操作バーの表示
    if (freeSlots.length > 0 || partialSlots.length > 0) {
      this.bulkActions.classList.remove('hidden');
      this.copyCheckedBtn.disabled = true;
      this.copyCheckedLabel.textContent = 'チェック済みをコピー';
    } else {
      this.bulkActions.classList.add('hidden');
    }

    if (freeSlots.length === 0) {
      this.freeSlots.innerHTML = `
        <div class="no-results">
          <p>被りなしの空き時間は見つかりませんでした。</p>
          <p style="font-size: 12px; margin-top: 4px;">「追加検索する」ボタンで被りありの候補を表示できます。</p>
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
        if (type === 'partial' || type === 'additional') {
          const names = slot.conflictingEmails.map((e) => e.split('@')[0]).join(', ');
          const overlapInfo = slot.overlapCount > slot.conflictCount
            ? `${slot.conflictCount}人被り・${slot.overlapCount}件`
            : `${slot.conflictCount}人被り`;
          badge = `<span class="slot-conflict-badge">${overlapInfo} (${this.escapeHtml(names)})</span>`;
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
    // コピーボタンの状態更新
    if (this.copyCheckedBtn) {
      this.copyCheckedBtn.disabled = checked.length === 0;
      this.copyCheckedLabel.textContent = checked.length > 0
        ? `チェック済みをコピー (${checked.length}件)`
        : 'チェック済みをコピー';
    }
  }

  // ============================================================
  // 一括選択・解除・コピー
  // ============================================================
  selectAllSlots() {
    document.querySelectorAll('.slot-checkbox').forEach((cb) => { cb.checked = true; });
    this.onSlotCheckChanged();
  }

  deselectAllSlots() {
    document.querySelectorAll('.slot-checkbox').forEach((cb) => { cb.checked = false; });
    this.onSlotCheckChanged();
  }

  copyCheckedSlots() {
    const checked = this.getCheckedSlots();
    if (checked.length === 0) return;

    const lines = checked.map((s) => {
      const start = new Date(s.start);
      const end = new Date(s.end);
      return `${this.formatDate(start)} ${this.formatTimeRange(start, end)}`;
    });
    const text = lines.join('\n');

    navigator.clipboard.writeText(text).then(() => {
      this.copyCheckedLabel.textContent = 'コピーしました!';
      setTimeout(() => {
        const count = this.getCheckedSlots().length;
        this.copyCheckedLabel.textContent = count > 0
          ? `チェック済みをコピー (${count}件)`
          : 'チェック済みをコピー';
      }, 1500);
    });
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
