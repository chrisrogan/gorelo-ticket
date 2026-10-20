/* taskpane.js — Full task pane controller.
   Handles the Create Ticket form and Settings view. */

// Extract the cache-busting ?v= param that taskpane.html injected into our URL.
// This ensures all sub-modules are also fetched with the same fresh timestamp.
const _v = new URL(import.meta.url).searchParams.get('v') || Date.now();

const { isAuthenticated, validateApiKey } = await import(`../auth/auth.js?v=${_v}`);
const {
  getClients, getContacts, getTechnicianGroups, getTechnicians,
  getTags, getEmailConfigurations, createTicket,
  PRIORITIES, SOURCES,
} = await import(`../gorelo-api.js?v=${_v}`);
const { getSetting, saveSettings, getAllSettings } = await import(`../settings.js?v=${_v}`);

// ─── State ────────────────────────────────────────────────────────────────────

let allTags = [];
let selectedTags = [];          // { id, name } objects; id may be null for new tags
let currentView = 'form';       // 'form' | 'settings' | 'setup'

// ─── Boot ─────────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  try {
    if (!isAuthenticated()) {
      showView('setup');
    } else {
      showView('form');
      await Promise.all([loadFormDropdowns(), loadEmailData()]);
    }
  } catch (err) {
    console.error('[Gorelo] Init error', err);
    showView('form');
  }
});

// ─── View switching ───────────────────────────────────────────────────────────

function showView(name) {
  currentView = name;
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById('view-' + name)?.classList.add('active');
}

// ─── Email data ───────────────────────────────────────────────────────────────

async function loadEmailData() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  // From field
  const from = item.from || {};
  const fromDisplay = from.displayName
    ? `${from.displayName} <${from.emailAddress}>`
    : (from.emailAddress || '');
  setVal('field-from', fromDisplay);

  // Subject
  setVal('field-subject', item.subject || '');

  // Body
  item.body.getAsync(Office.CoercionType.Text, {}, result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // Prepend sender info so context is preserved
      const senderLine = fromDisplay ? `From: ${fromDisplay}\n\n` : '';
      setVal('field-notes', senderLine + (result.value || '').trim());
    }
  });
}

// ─── Form dropdowns ───────────────────────────────────────────────────────────

async function loadFormDropdowns() {
  const [groups, tags, mailboxes] = await Promise.allSettled([
    getTechnicianGroups(),
    getTags(),
    getEmailConfigurations(),
  ]);

  // Groups
  if (groups.status === 'fulfilled') {
    populateSelect('field-group', groups.value, 'id', 'name', '', '— Select group —');
    const defaultGroup = getSetting('defaultTechnicianGroupId');
    if (defaultGroup) {
      setVal('field-group', String(defaultGroup));
      onGroupChange(String(defaultGroup));  // load assignees for default group
    }
  }

  // Tags cache
  if (tags.status === 'fulfilled') {
    allTags = tags.value || [];
  }

  // Mailboxes
  if (mailboxes.status === 'fulfilled') {
    populateSelect('field-mailbox', mailboxes.value, 'id', 'name', '', '— None —');
    const defaultMailbox = getSetting('defaultEmailConfigurationId');
    if (defaultMailbox) setVal('field-mailbox', String(defaultMailbox));
  }

  // Priorities / Sources (static)
  populateSelect('field-priority', PRIORITIES, 'id', 'name');
  setVal('field-priority', String(getSetting('defaultPriorityId') || 2));

  populateSelect('field-source', SOURCES, 'id', 'name');
  setVal('field-source', String(getSetting('defaultSourceId') || 2));

  // Load clients
  await loadClients();
}

async function loadClients() {
  try {
    const clients = await getClients();
    populateSelect('field-client', clients, 'id', 'name', '', '— Select client —');

    // Auto-match by sender email
    const fromEmail = Office.context.mailbox.item?.from?.emailAddress || '';
    if (fromEmail) {
      await tryAutoMatchClient(fromEmail, clients);
    }
  } catch (err) {
    console.warn('[Gorelo] Could not load clients', err.message);
  }
}

async function tryAutoMatchClient(email, clients) {
  // Try to find a contact with matching email, which tells us the client
  try {
    const contacts = await getContacts();
    const match = contacts.find(c =>
      (c.email || '').toLowerCase() === email.toLowerCase()
    );
    if (match?.clientId) {
      setVal('field-client', String(match.clientId));
      await onClientChange(match.clientId, match.id);
    }
  } catch { /* non-fatal */ }
}

async function onClientChange(clientId, preselectContactId) {
  const contactSel = document.getElementById('field-contact');
  contactSel.innerHTML = '<option value="">— Loading… —</option>';

  try {
    const contacts = await getContacts(clientId || null);
    console.log('[Gorelo] getContacts(' + clientId + ') returned', contacts.length, 'contacts:', contacts.slice(0,3));
    populateSelect('field-contact', contacts, 'id', 'name', '', '— Select contact —');
    if (preselectContactId) setVal('field-contact', String(preselectContactId));
  } catch (err) {
    console.error('[Gorelo] getContacts error:', err);
    contactSel.innerHTML = '<option value="">— Select contact —</option>';
  }
}

async function onGroupChange(groupId) {
  const assigneeSel = document.getElementById('field-assignee');
  assigneeSel.innerHTML = '<option value="">— Loading… —</option>';

  try {
    const techs = await getTechnicians(groupId || null);
    console.log('[Gorelo] getTechnicians(' + groupId + ') returned', techs.length, 'techs:', techs.slice(0,3));
    populateSelect('field-assignee', techs, 'id', 'name', '', '— Unassigned —');

    const defaultAssignee = getSetting('defaultAssignedToId');
    if (defaultAssignee) setVal('field-assignee', String(defaultAssignee));
  } catch (err) {
    console.error('[Gorelo] getTechnicians error:', err);
    assigneeSel.innerHTML = '<option value="">— Unassigned —</option>';
  }
}

// ─── Tag input ────────────────────────────────────────────────────────────────

function initTagInput() {
  const input = document.getElementById('tag-text');
  const suggestions = document.getElementById('tag-suggestions');

  input.addEventListener('input', () => {
    const q = input.value.trim().toLowerCase();
    if (!q) { suggestions.classList.remove('visible'); return; }

    const matches = allTags.filter(t =>
      (t.name || t.tagName || '').toLowerCase().includes(q) &&
      !selectedTags.find(s => s.id === (t.tagId || t.id))
    ).slice(0, 8);

    suggestions.innerHTML = '';
    matches.forEach(t => {
      const name = t.name || t.tagName;
      const id = t.tagId || t.id;
      const div = document.createElement('div');
      div.className = 'tag-suggestion-item';
      div.textContent = name;
      div.addEventListener('mousedown', e => {
        e.preventDefault();
        addTag({ id, name });
        input.value = '';
        suggestions.classList.remove('visible');
      });
      suggestions.appendChild(div);
    });

    if (matches.length > 0) {
      suggestions.classList.add('visible');
    } else {
      suggestions.classList.remove('visible');
    }
  });

  input.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ',') {
      e.preventDefault();
      const val = input.value.trim().replace(/,$/, '');
      if (val) {
        // New tag (no ID)
        addTag({ id: null, name: val });
        input.value = '';
        suggestions.classList.remove('visible');
      }
    }
    if (e.key === 'Backspace' && !input.value && selectedTags.length) {
      selectedTags.pop();
      renderTags();
    }
  });

  input.addEventListener('blur', () => {
    setTimeout(() => suggestions.classList.remove('visible'), 150);
  });
}

function addTag(tag) {
  if (!selectedTags.find(t => t.name.toLowerCase() === tag.name.toLowerCase())) {
    selectedTags.push(tag);
    renderTags();
  }
}

function renderTags() {
  const wrap = document.getElementById('tag-wrap');
  const input = document.getElementById('tag-text');
  // Remove existing chips
  wrap.querySelectorAll('.tag-chip').forEach(c => c.remove());
  // Re-insert chips before the input
  selectedTags.forEach((tag, i) => {
    const chip = document.createElement('span');
    chip.className = 'tag-chip';
    chip.innerHTML = `${tag.name} <span class="remove" data-idx="${i}">×</span>`;
    chip.querySelector('.remove').addEventListener('click', () => {
      selectedTags.splice(i, 1);
      renderTags();
    });
    wrap.insertBefore(chip, input);
  });
}

// ─── Submit ───────────────────────────────────────────────────────────────────

async function submitTicket() {
  console.log('[Gorelo] submitTicket called, subject:', getVal('field-subject'), 'groupId:', getVal('field-group'));
  const subject = getVal('field-subject').trim();
  const groupId = getVal('field-group');

  if (!subject) { showFormAlert('Please enter a ticket title.', 'error'); scrollToAlert(); return; }
  if (!groupId) { showFormAlert('Please select a group (required by Gorelo).', 'error'); scrollToAlert(); return; }

  setSubmitting(true);
  hideFormAlert();

  try {
    const result = await createTicket({
      title: subject,
      description: getVal('field-notes'),
      markdownComment: getVal('field-notes'),
      clientId: intVal('field-client'),
      clientName: getVal('field-client') ? document.getElementById('field-client').selectedOptions[0]?.text : '',
      contactId: intVal('field-contact'),
      outboundContactEmail: Office.context.mailbox.item?.from?.emailAddress || '',
      technicianGroupIds: groupId ? [Number(groupId)] : [],
      assignedToId: intVal('field-assignee'),
      priorityId: intVal('field-priority') || 2,
      sourceId: intVal('field-source') || 2,
      emailConfigurationId: intVal('field-mailbox') || null,
      newTags: selectedTags.filter(t => t.id === null).map(t => t.name),
      existingTagIds: selectedTags.filter(t => t.id !== null).map(t => String(t.id)),
      sendNotification: document.getElementById('toggle-notify').checked,
    });

    const ticketNum = result?.ticketNumberWithPrefix || result?.ticketNumber || '';
    showFormAlert(
      ticketNum ? `✓ Ticket ${ticketNum} created successfully.` : '✓ Ticket created successfully.',
      'success'
    );

    // Reset form for next use
    setTimeout(() => {
      selectedTags = [];
      renderTags();
      hideFormAlert();
    }, 4000);

  } catch (err) {
    showFormAlert('Error: ' + (err.message || 'Could not create ticket.'), 'error');
  } finally {
    setSubmitting(false);
  }
}

// ─── Settings ─────────────────────────────────────────────────────────────────

async function loadSettingsView() {
  const settings = getAllSettings();

  // API key
  document.getElementById('settings-api-key').value = settings.apiKey || '';

  // Load dropdowns for defaults
  try {
    const [groups, techs, mailboxes] = await Promise.allSettled([
      getTechnicianGroups(),
      getTechnicians(),
      getEmailConfigurations(),
    ]);

    if (groups.status === 'fulfilled') {
      populateSelect('settings-default-group', groups.value, 'id', 'name', settings.defaultTechnicianGroupId || '', '— None —');
    }
    if (techs.status === 'fulfilled') {
      populateSelect('settings-default-assignee', techs.value, 'id', 'name', settings.defaultAssignedToId || '', '— None —');
    }
    if (mailboxes.status === 'fulfilled') {
      populateSelect('settings-default-mailbox', mailboxes.value, 'id', 'name', settings.defaultEmailConfigurationId || '', '— None —');
    }
  } catch { /* non-fatal */ }

  populateSelect('settings-default-priority', PRIORITIES, 'id', 'name', settings.defaultPriorityId || 2);
  populateSelect('settings-default-source', SOURCES, 'id', 'name', settings.defaultSourceId || 2);

  document.getElementById('settings-auto-open').checked = !!settings.autoOpenTaskPane;
}

async function saveSettingsFromUI() {
  hideSettingsAlert();

  try {
    await saveSettings({
      defaultTechnicianGroupId: getVal('settings-default-group'),
      defaultAssignedToId: getVal('settings-default-assignee'),
      defaultPriorityId: intVal('settings-default-priority') || 2,
      defaultSourceId: intVal('settings-default-source') || 2,
      defaultEmailConfigurationId: getVal('settings-default-mailbox'),
      autoOpenTaskPane: document.getElementById('settings-auto-open').checked,
    });
    showSettingsAlert('✓ Defaults saved.', 'success');
    setTimeout(hideSettingsAlert, 2500);
  } catch (err) {
    showSettingsAlert('Error saving: ' + err.message, 'error');
  }
}

async function validateAndSaveApiKey() {
  const key = document.getElementById('settings-api-key').value.trim();
  const statusEl = document.getElementById('api-key-status');

  statusEl.textContent = 'Validating…';
  statusEl.className = 'api-key-status';
  statusEl.classList.remove('hidden');

  const result = await validateApiKey(key);

  if (result.valid) {
    statusEl.textContent = '✓ API key is valid';
    statusEl.classList.add('ok');

    await saveSettings({ apiKey: key });
    showSettingsAlert('✓ API key saved successfully.', 'success');

    // If we were on the setup screen, load the form now
    if (currentView === 'settings') {
      // re-enable form for next navigation
      setTimeout(() => showSettingsAlert('', 'hidden'), 3000);
    }
  } else {
    statusEl.textContent = '✗ ' + result.error;
    statusEl.classList.add('err');
  }
}

// ─── Event wiring ─────────────────────────────────────────────────────────────

// DOM is already ready (module loaded dynamically after DOMContentLoaded)
;(() => {
  // Header buttons
  document.getElementById('btn-settings').addEventListener('click', async () => {
    showView('settings');
    await loadSettingsView();
  });
  document.getElementById('btn-refresh').addEventListener('click', async () => {
    if (currentView === 'form') {
      await loadEmailData();
    }
  });

  // Setup view
  document.getElementById('btn-goto-settings').addEventListener('click', async () => {
    showView('settings');
    await loadSettingsView();
  });

  // Form
  document.getElementById('field-client').addEventListener('change', e => {
    onClientChange(e.target.value);
  });
  document.getElementById('field-group').addEventListener('change', e => {
    onGroupChange(e.target.value);
  });
  const btnSubmit = document.getElementById('btn-submit');
  console.log('[Gorelo] btn-submit element:', btnSubmit);
  if (btnSubmit) btnSubmit.addEventListener('click', submitTicket);
  else console.error('[Gorelo] btn-submit NOT FOUND in DOM');

  // Settings
  document.getElementById('btn-validate-key').addEventListener('click', validateAndSaveApiKey);
  document.getElementById('btn-save-settings').addEventListener('click', saveSettingsFromUI);
  document.getElementById('btn-back').addEventListener('click', async () => {
    if (isAuthenticated()) {
      showView('form');
    } else {
      showView('setup');
    }
  });
  document.getElementById('btn-show-key').addEventListener('click', () => {
    const input = document.getElementById('settings-api-key');
    const btn = document.getElementById('btn-show-key');
    if (input.type === 'password') {
      input.type = 'text';
      btn.textContent = 'Hide';
    } else {
      input.type = 'password';
      btn.textContent = 'Show';
    }
  });

  initTagInput();
});

// ─── Utilities ────────────────────────────────────────────────────────────────

function populateSelect(id, items, valueKey, labelKey, selectedValue, emptyLabel) {
  const sel = document.getElementById(id);
  if (!sel) return;
  const current = selectedValue !== undefined ? String(selectedValue) : sel.value;
  sel.innerHTML = emptyLabel ? `<option value="">${emptyLabel}</option>` : '';
  (items || []).forEach(item => {
    const val = String(item[valueKey] ?? '');
    const label = typeof labelKey === 'function' ? labelKey(item) : (item[labelKey] ?? val);
    const opt = new Option(label, val);
    if (val && val === current) opt.selected = true;
    sel.add(opt);
  });
}

function getVal(id) {
  return document.getElementById(id)?.value || '';
}
function setVal(id, val) {
  const el = document.getElementById(id);
  if (el) el.value = val;
}
function intVal(id) {
  const v = getVal(id);
  const n = parseInt(v, 10);
  return isNaN(n) ? 0 : n;
}

function setSubmitting(on) {
  const btn = document.getElementById('btn-submit');
  const text = document.getElementById('btn-submit-text');
  const spinner = document.getElementById('btn-submit-spinner');
  btn.disabled = on;
  text.textContent = on ? 'Creating…' : 'Create Ticket';
  spinner.classList.toggle('hidden', !on);
}

function scrollToAlert() {
  document.getElementById('form-alert')?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function showFormAlert(msg, type) {
  const el = document.getElementById('form-alert');
  el.textContent = msg;
  el.className = `alert alert-${type}`;
  el.classList.remove('hidden');
}
function hideFormAlert() {
  document.getElementById('form-alert').classList.add('hidden');
}
function showSettingsAlert(msg, type) {
  const el = document.getElementById('settings-alert');
  if (!msg) { el.classList.add('hidden'); return; }
  el.textContent = msg;
  el.className = `alert alert-${type}`;
  el.classList.remove('hidden');
}
function hideSettingsAlert() {
  document.getElementById('settings-alert').classList.add('hidden');
}
