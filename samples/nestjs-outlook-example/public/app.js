/**
 * NestJS Outlook Demo Application
 * Client-side JavaScript for interacting with the demo API endpoints
 */

(function() {
  'use strict';

  // API base URL - adjust if running on a different port
  const API_BASE = '';

  /**
   * Initialize the application
   */
  function init() {
    setupTabs();
    setupEventListeners();
    setDefaultDateTimes();
    setupConsentMessageListener();
  }

  /**
   * Set up tab navigation
   */
  function setupTabs() {
    const tabs = document.querySelectorAll('.tab');
    const panels = document.querySelectorAll('.tab-panel');

    tabs.forEach(tab => {
      tab.addEventListener('click', () => {
        const targetId = `tab-${tab.dataset.tab}`;

        // Update tab states
        tabs.forEach(t => t.classList.remove('tab--active'));
        tab.classList.add('tab--active');

        // Update panel states
        panels.forEach(panel => {
          panel.classList.remove('tab-panel--active');
          if (panel.id === targetId) {
            panel.classList.add('tab-panel--active');
          }
        });
      });
    });
  }

  /**
   * Set up all event listeners
   */
  function setupEventListeners() {
    // Authentication
    document.getElementById('btnUserLogin').addEventListener('click', handleUserLogin);

    // Tenant onboarding wizard
    document.getElementById('modeShared').addEventListener('click', () => setMode('shared'));
    document.getElementById('modeDedicated').addEventListener('click', () => setMode('dedicated'));
    document.getElementById('btnGenerateCert').addEventListener('click', handleGenerateCert);
    document.getElementById('btnDownloadCer').addEventListener('click', handleDownloadCer);
    document.getElementById('btnRegisterConsent').addEventListener('click', handleRegisterConsent);

    // Calendar
    document.getElementById('createEventForm').addEventListener('submit', handleCreateEvent);
    document.getElementById('btnGetTenantCalendars').addEventListener('click', handleGetTenantCalendars);
    document.getElementById('btnGetTenantEvents').addEventListener('click', handleGetTenantEvents);

    // Email
    document.getElementById('sendEmailForm').addEventListener('submit', handleSendEmail);
    document.getElementById('btnGetEmails').addEventListener('click', handleGetEmails);

    // Users
    document.getElementById('btnListUsers').addEventListener('click', handleListUsers);
    document.getElementById('btnGetUser').addEventListener('click', handleGetUser);
  }

  /**
   * Set default date/time values for calendar form
   */
  function setDefaultDateTimes() {
    const now = new Date();
    const later = new Date(now.getTime() + 60 * 60 * 1000); // 1 hour later

    document.getElementById('eventStart').value = formatDateTimeLocal(now);
    document.getElementById('eventEnd').value = formatDateTimeLocal(later);
  }

  /**
   * Format date for datetime-local input
   */
  function formatDateTimeLocal(date) {
    const pad = n => n.toString().padStart(2, '0');
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
  }

  /**
   * Listen for messages from consent popup window
   */
  function setupConsentMessageListener() {
    window.addEventListener('message', (event) => {
      if (event.data && event.data.type) {
        switch (event.data.type) {
          case 'tenant-consent-success':
            showResponse('adminConsentResponse', 'adminConsentStatusCode', 'adminConsentJson', {
              success: true,
              message: 'Admin consent granted successfully!',
              tenantId: event.data.tenantId
            }, 200);
            updateConnectionStatus(true, 'Tenant Connected');
            setConsentStatus('Tenant connected — admin consent granted and access verified.', true);
            break;
          case 'tenant-consent-failed':
          case 'tenant-consent-error':
            showResponse('adminConsentResponse', 'adminConsentStatusCode', 'adminConsentJson', {
              success: false,
              error: event.data.error || 'Consent failed'
            }, 400);
            setConsentStatus(event.data.error || 'Admin consent failed.', false);
            break;
        }
      }
    });
  }

  /**
   * Update the connection status in the header
   */
  function updateConnectionStatus(connected, text) {
    const statusEl = document.getElementById('connectionStatus');
    const indicator = statusEl.querySelector('.status-indicator');
    const textEl = statusEl.querySelector('span:not(.status-indicator)');

    if (connected) {
      statusEl.classList.remove('app-header__status--disconnected');
      statusEl.classList.add('app-header__status--connected');
      indicator.classList.remove('status-indicator--disconnected');
      indicator.classList.add('status-indicator--connected');
    } else {
      statusEl.classList.remove('app-header__status--connected');
      statusEl.classList.add('app-header__status--disconnected');
      indicator.classList.remove('status-indicator--connected');
      indicator.classList.add('status-indicator--disconnected');
    }

    textEl.textContent = text || (connected ? 'Connected' : 'Not Connected');
  }

  /**
   * Make an API request
   */
  async function apiRequest(method, endpoint, data = null) {
    const options = {
      method,
      headers: {
        'Content-Type': 'application/json',
      },
    };

    if (data && method !== 'GET') {
      options.body = JSON.stringify(data);
    }

    const response = await fetch(`${API_BASE}${endpoint}`, options);
    const json = await response.json().catch(() => ({}));

    return {
      status: response.status,
      ok: response.ok,
      data: json,
    };
  }

  /**
   * Show response in a panel
   */
  function showResponse(panelId, statusCodeId, jsonId, data, status) {
    const panel = document.getElementById(panelId);
    const statusCode = document.getElementById(statusCodeId);
    const json = document.getElementById(jsonId);

    panel.classList.remove('hidden');
    statusCode.textContent = status;
    statusCode.className = 'response-panel__status-code';
    statusCode.classList.add(status >= 200 && status < 300 ? 'response-panel__status-code--success' : 'response-panel__status-code--error');
    json.textContent = JSON.stringify(data, null, 2);
  }

  /**
   * Set button loading state
   */
  function setButtonLoading(button, loading) {
    if (loading) {
      button.disabled = true;
      button.dataset.originalText = button.innerHTML;
      button.innerHTML = '<span class="spinner"></span> Loading...';
    } else {
      button.disabled = false;
      button.innerHTML = button.dataset.originalText;
    }
  }

  // ============================================
  // Authentication Handlers
  // ============================================

  /**
   * Handle user OAuth login
   */
  async function handleUserLogin() {
    const button = document.getElementById('btnUserLogin');
    setButtonLoading(button, true);

    try {
      const result = await apiRequest('GET', '/auth/microsoft/login');
      showResponse('userAuthResponse', 'userAuthStatusCode', 'userAuthJson', result.data, result.status);

      if (result.ok && result.data.url) {
        // Redirect to Microsoft login
        window.location.href = result.data.url;
      }
    } catch (error) {
      showResponse('userAuthResponse', 'userAuthStatusCode', 'userAuthJson', { error: error.message }, 500);
    } finally {
      setButtonLoading(button, false);
    }
  }

  // ============================================
  // Tenant Onboarding Wizard
  // ============================================

  // Current onboarding mode and the most recently generated certificate.
  let currentMode = 'shared';
  let generatedCert = null;

  /**
   * Switch between the "shared" and "dedicated" onboarding models.
   */
  function setMode(mode) {
    currentMode = mode;
    const dedicated = mode === 'dedicated';

    document.getElementById('modeShared').classList.toggle('mode-toggle__btn--active', !dedicated);
    document.getElementById('modeDedicated').classList.toggle('mode-toggle__btn--active', dedicated);
    document.getElementById('clientIdGroup').classList.toggle('hidden', !dedicated);
    document.getElementById('generateCertGroup').classList.toggle('hidden', !dedicated);

    document.getElementById('modeHint').innerHTML = dedicated
      ? 'Uses <strong>your own</strong> Azure app registration. Generate a certificate below, upload it to your app, then grant consent.'
      : 'Uses the Checkfirst-owned app &amp; certificate. Just enter your Microsoft Tenant ID and grant consent — no certificate needed.';

    // A mode switch invalidates any previously generated certificate.
    generatedCert = null;
    document.getElementById('certOutput').classList.add('hidden');
  }

  /**
   * Generate a dedicated per-tenant certificate.
   */
  async function handleGenerateCert() {
    const button = document.getElementById('btnGenerateCert');
    const tenantId = document.getElementById('microsoftTenantId').value.trim();

    if (!tenantId) {
      setConsentStatus('Enter your Microsoft Tenant ID first.', false);
      return;
    }

    setButtonLoading(button, true);
    try {
      const result = await apiRequest('POST', '/tenant/certificate/generate', { tenantId });
      showResponse('adminConsentResponse', 'adminConsentStatusCode', 'adminConsentJson', result.data, result.status);

      if (result.ok) {
        generatedCert = result.data;
        document.getElementById('certThumbprint').textContent = result.data.thumbprint;
        document.getElementById('certOutput').classList.remove('hidden');
        setConsentStatus('Certificate generated. Upload the .cer to your Azure app, then grant consent.', true);
      } else {
        setConsentStatus('Certificate generation failed.', false);
      }
    } catch (error) {
      showResponse('adminConsentResponse', 'adminConsentStatusCode', 'adminConsentJson', { error: error.message }, 500);
    } finally {
      setButtonLoading(button, false);
    }
  }

  /**
   * Download the generated public certificate as a .cer file.
   */
  function handleDownloadCer() {
    if (!generatedCert || !generatedCert.certificatePem) return;
    const tenantId = document.getElementById('microsoftTenantId').value.trim() || 'tenant';
    const blob = new Blob([generatedCert.certificatePem], { type: 'application/x-x509-ca-cert' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${tenantId}.cer`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  /**
   * Register the tenant and open the admin-consent popup.
   */
  async function handleRegisterConsent() {
    const button = document.getElementById('btnRegisterConsent');
    const tenantId = document.getElementById('microsoftTenantId').value.trim();

    if (!tenantId) {
      setConsentStatus('Enter your Microsoft Tenant ID.', false);
      return;
    }

    let body;
    if (currentMode === 'dedicated') {
      const clientId = document.getElementById('tenantClientId').value.trim();
      if (!clientId) {
        setConsentStatus('Enter your Application (client) ID.', false);
        return;
      }
      if (!generatedCert) {
        setConsentStatus('Generate a certificate first.', false);
        return;
      }
      body = {
        mode: 'dedicated',
        tenantId,
        clientId,
        certificateThumbprint: generatedCert.thumbprint,
        certificateKeyPath: generatedCert.keyPath,
        certificatePath: generatedCert.certPath,
      };
    } else {
      body = { mode: 'shared', tenantId };
    }

    setButtonLoading(button, true);
    try {
      const result = await apiRequest('POST', '/tenant/register', body);
      showResponse('adminConsentResponse', 'adminConsentStatusCode', 'adminConsentJson', result.data, result.status);

      if (result.ok && result.data.adminConsentUrl) {
        setConsentStatus('Tenant registered. Complete admin consent in the popup window…', true);
        openConsentPopup(result.data.adminConsentUrl);
      } else {
        setConsentStatus('Registration failed — see the response below.', false);
      }
    } catch (error) {
      showResponse('adminConsentResponse', 'adminConsentStatusCode', 'adminConsentJson', { error: error.message }, 500);
    } finally {
      setButtonLoading(button, false);
    }
  }

  /**
   * Open the Microsoft admin-consent URL in a centered popup.
   */
  function openConsentPopup(url) {
    const width = 600;
    const height = 700;
    const left = (window.innerWidth - width) / 2 + window.screenX;
    const top = (window.innerHeight - height) / 2 + window.screenY;
    window.open(
      url,
      'AdminConsent',
      `width=${width},height=${height},left=${left},top=${top},toolbar=no,menubar=no,scrollbars=yes`
    );
  }

  /**
   * Show a human-readable status line under the wizard.
   */
  function setConsentStatus(message, ok) {
    const el = document.getElementById('consentStatus');
    el.textContent = message;
    el.classList.remove('hidden', 'consent-status--ok', 'consent-status--error');
    el.classList.add(ok ? 'consent-status--ok' : 'consent-status--error');
  }

  // ============================================
  // Calendar Handlers
  // ============================================

  /**
   * Handle create event form submission
   */
  async function handleCreateEvent(event) {
    event.preventDefault();

    const button = document.getElementById('btnCreateEvent');
    const name = document.getElementById('eventName').value;
    const startDateTime = document.getElementById('eventStart').value;
    const endDateTime = document.getElementById('eventEnd').value;

    setButtonLoading(button, true);

    try {
      const result = await apiRequest('POST', '/calendar/events', {
        name,
        startDateTime: new Date(startDateTime).toISOString(),
        endDateTime: new Date(endDateTime).toISOString(),
      });

      showResponse('createEventResponse', 'createEventStatusCode', 'createEventJson', result.data, result.status);

      if (result.ok) {
        // Reset form on success
        document.getElementById('eventName').value = '';
        setDefaultDateTimes();
      }
    } catch (error) {
      showResponse('createEventResponse', 'createEventStatusCode', 'createEventJson', { error: error.message }, 500);
    } finally {
      setButtonLoading(button, false);
    }
  }

  /**
   * Handle get tenant calendars
   */
  async function handleGetTenantCalendars() {
    const button = document.getElementById('btnGetTenantCalendars');
    const userId = document.getElementById('tenantCalendarUserId').value;

    if (!userId) {
      showResponse('tenantCalendarResponse', 'tenantCalendarStatusCode', 'tenantCalendarJson',
        { error: 'User ID or email is required' }, 400);
      return;
    }

    setButtonLoading(button, true);

    try {
      const result = await apiRequest('GET', `/tenant/users/${encodeURIComponent(userId)}/calendars`);
      showResponse('tenantCalendarResponse', 'tenantCalendarStatusCode', 'tenantCalendarJson', result.data, result.status);
    } catch (error) {
      showResponse('tenantCalendarResponse', 'tenantCalendarStatusCode', 'tenantCalendarJson', { error: error.message }, 500);
    } finally {
      setButtonLoading(button, false);
    }
  }

  /**
   * Handle get tenant calendar events
   */
  async function handleGetTenantEvents() {
    const button = document.getElementById('btnGetTenantEvents');
    const userId = document.getElementById('tenantCalendarUserId').value;

    if (!userId) {
      showResponse('tenantCalendarResponse', 'tenantCalendarStatusCode', 'tenantCalendarJson',
        { error: 'User ID or email is required' }, 400);
      return;
    }

    setButtonLoading(button, true);

    try {
      const result = await apiRequest('GET', `/tenant/users/${encodeURIComponent(userId)}/events`);
      showResponse('tenantCalendarResponse', 'tenantCalendarStatusCode', 'tenantCalendarJson', result.data, result.status);
    } catch (error) {
      showResponse('tenantCalendarResponse', 'tenantCalendarStatusCode', 'tenantCalendarJson', { error: error.message }, 500);
    } finally {
      setButtonLoading(button, false);
    }
  }

  // ============================================
  // Email Handlers
  // ============================================

  /**
   * Handle send email form submission
   */
  async function handleSendEmail(event) {
    event.preventDefault();

    const button = document.getElementById('btnSendEmail');
    const to = document.getElementById('emailTo').value;
    const subject = document.getElementById('emailSubject').value;
    const body = document.getElementById('emailBody').value;

    setButtonLoading(button, true);

    try {
      const result = await apiRequest('POST', '/email/send', {
        to,
        subject,
        body,
      });

      showResponse('sendEmailResponse', 'sendEmailStatusCode', 'sendEmailJson', result.data, result.status);

      if (result.ok) {
        // Reset form on success
        document.getElementById('emailTo').value = '';
        document.getElementById('emailSubject').value = '';
        document.getElementById('emailBody').value = '';
      }
    } catch (error) {
      showResponse('sendEmailResponse', 'sendEmailStatusCode', 'sendEmailJson', { error: error.message }, 500);
    } finally {
      setButtonLoading(button, false);
    }
  }

  /**
   * Handle get emails
   */
  async function handleGetEmails() {
    const button = document.getElementById('btnGetEmails');
    const limit = document.getElementById('emailLimit').value;

    setButtonLoading(button, true);

    try {
      const result = await apiRequest('GET', `/email/inbox?limit=${limit}`);
      showResponse('getEmailsResponse', 'getEmailsStatusCode', 'getEmailsJson', result.data, result.status);
    } catch (error) {
      showResponse('getEmailsResponse', 'getEmailsStatusCode', 'getEmailsJson', { error: error.message }, 500);
    } finally {
      setButtonLoading(button, false);
    }
  }

  // ============================================
  // Users Handlers
  // ============================================

  /**
   * Handle list users
   */
  async function handleListUsers() {
    const button = document.getElementById('btnListUsers');
    const filter = document.getElementById('userSearchFilter').value;

    setButtonLoading(button, true);

    try {
      const params = filter ? `?filter=${encodeURIComponent(filter)}` : '';
      const result = await apiRequest('GET', `/tenant/users${params}`);
      showResponse('listUsersResponse', 'listUsersStatusCode', 'listUsersJson', result.data, result.status);
    } catch (error) {
      showResponse('listUsersResponse', 'listUsersStatusCode', 'listUsersJson', { error: error.message }, 500);
    } finally {
      setButtonLoading(button, false);
    }
  }

  /**
   * Handle get user details
   */
  async function handleGetUser() {
    const button = document.getElementById('btnGetUser');
    const userId = document.getElementById('userDetailId').value;

    if (!userId) {
      showResponse('getUserResponse', 'getUserStatusCode', 'getUserJson',
        { error: 'User ID or email is required' }, 400);
      return;
    }

    setButtonLoading(button, true);

    try {
      const result = await apiRequest('GET', `/tenant/users/${encodeURIComponent(userId)}`);
      showResponse('getUserResponse', 'getUserStatusCode', 'getUserJson', result.data, result.status);
    } catch (error) {
      showResponse('getUserResponse', 'getUserStatusCode', 'getUserJson', { error: error.message }, 500);
    } finally {
      setButtonLoading(button, false);
    }
  }

  // Initialize when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
