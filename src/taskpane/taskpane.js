/* taskpane.js — Full task pane controller.
   Handles the Create Ticket form and Settings view. */


// Extract the cache-busting ?v= param that taskpane.html injected into our URL.
// This ensures all sub-modules are also fetched with the same fresh timestamp.
const _v = new URL(import.meta.url).searchParams.get('v') || Date.now();

const { isAuthenticated, validateApiKey } = await import(`../auth/auth.js?v=${_v}`);
const {
  getClients, getContacts, searchContacts, createContact,
  getTechnicianGroups, getTechnicians,
  getTicketTypes, getTags, getEmailConfigurations, createTicket,
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
  const [groups, types, tags, mailboxes] = await Promise.allSettled([
    getTechnicianGroups(),
    getTicketTypes(),
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

  if (types.status === 'fulfilled' && types.value.length > 0) {
    populateSelect('field-type', types.value, 'id', 'name', '', '— Select type —');
    const defaultType = getSetting('defaultTypeId');
    if (defaultType) setVal('field-type', String(defaultType));
    else setVal('field-type', String(types.value[0].id)); // default to first type
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

  // Apply notify default (off unless explicitly enabled in settings)
  document.getElementById('toggle-notify').checked = !!getSetting('defaultSendNotification');

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
  // Search contacts by sender email — fast targeted lookup
  try {
    const matches = await searchContacts(email);
    const match = matches.find(c => (c.email || '').toLowerCase() === email.toLowerCase());
    if (match?.clientId) {
      setVal('field-client', String(match.clientId));
      await onClientChange(match.clientId, match.id);
    } else if (match) {
      // Contact exists but is temporary (no client yet) — still pre-select them
      await onClientChange(null, match.id);
    }
    // Store the match result so the contact section can use it
    window._senderContact = match || null;
  } catch {
    window._senderContact = null;
  }
  updateCreateContactButton(email);
}

function updateCreateContactButton(senderEmail) {
  const btn = document.getElementById('btn-create-contact');
  const hint = document.getElementById('create-contact-hint');
  if (!btn) return;
  if (!senderEmail) { btn.style.display = 'none'; return; }
  if (window._senderContact) {
    // Already exists — hide button, show hint
    btn.style.display = 'none';
    if (hint) { hint.textContent = '✓ Contact found in Gorelo'; hint.style.display = 'block'; }
  } else {
    // Not found — show create button
    btn.style.display = 'inline-block';
    if (hint) hint.style.display = 'none';
  }
}

async function handleCreateContact() {
  const btn = document.getElementById('btn-create-contact');
  const hint = document.getElementById('create-contact-hint');
  const fromRaw = getVal('field-from');       // "Display Name <email@example.com>"
  const clientId = intVal('field-client');

  // Parse display name and email from the From field
  const emailMatch = fromRaw.match(/<([^>]+)>/);
  const email = emailMatch ? emailMatch[1].trim() : fromRaw.trim();
  const displayName = emailMatch ? fromRaw.replace(/<[^>]+>/, '').trim() : '';
  const nameParts = displayName.split(/\s+/);
  const firstName = nameParts[0] || email.split('@')[0];
  const lastName  = nameParts.slice(1).join(' ') || '—';

  if (!email) { showError('Cannot determine sender email address.'); return; }

  btn.disabled = true;
  btn.textContent = 'Creating…';
  try {
    const contact = await createContact({ firstName, lastName, email, clientId: clientId || null });
    window._senderContact = contact;
    // Refresh the contact dropdown and select the new contact
    await onClientChange(clientId || null, contact.id);
    btn.style.display = 'none';
    if (hint) { hint.textContent = '✓ Contact created in Gorelo'; hint.style.display = 'block'; }
  } catch (err) {
    btn.disabled = false;
    btn.textContent = '+ Create contact';
    showError('Could not create contact: ' + err.message);
  }
}

async function onClientChange(clientId, preselectContactId) {
  const contactSel = document.getElementById('field-contact');
  contactSel.innerHTML = '<option value="">— Loading… —</option>';

  try {
    const contacts = await getContacts(clientId || null);
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
  const subject = getVal('field-subject').trim();
  const groupId = getVal('field-group');

  if (!subject) { showFormAlert('Please enter a ticket title.', 'error'); scrollToAlert(); return; }
  if (!groupId) { showFormAlert('Please select a group (required by Gorelo).', 'error'); scrollToAlert(); return; }

  setSubmitting(true);
  hideFormAlert();

  try {
    const groupIdNum = groupId ? Number(groupId) : null;
    const result = await createTicket({
      title: subject,
      description: getVal('field-notes'),
      clientId: intVal('field-client'),
      contactId: intVal('field-contact'),
      technicianGroupIds: (groupIdNum && groupIdNum > 0) ? [groupIdNum] : [],
      assignedToId: intVal('field-assignee'),
      typeId: intVal('field-type') || null,
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
    const [groups, techs, mailboxes, types] = await Promise.allSettled([
      getTechnicianGroups(),
      getTechnicians(),
      getEmailConfigurations(),
      getTicketTypes(),
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
    if (types.status === 'fulfilled') {
      populateSelect('settings-default-type', types.value, 'id', 'name', settings.defaultTypeId || '', '— None —');
    }
  } catch { /* non-fatal */ }

  populateSelect('settings-default-priority', PRIORITIES, 'id', 'name', settings.defaultPriorityId || 2);
  populateSelect('settings-default-source', SOURCES, 'id', 'name', settings.defaultSourceId || 2);

  document.getElementById('settings-auto-open').checked = !!settings.autoOpenTaskPane;
  document.getElementById('settings-default-notify').checked = !!settings.defaultSendNotification;
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
      defaultTypeId: intVal('settings-default-type') || null,
      defaultSendNotification: document.getElementById('settings-default-notify').checked,
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
  function on(id, event, fn) {
    const el = document.getElementById(id);
    if (el) el.addEventListener(event, fn);
    else console.warn('[Gorelo] element not found:', id);
  }

  // Header buttons
  on('btn-settings', 'click', async () => {
    showView('settings');
    await loadSettingsView();
  });
  on('btn-refresh', 'click', async () => {
    if (currentView === 'form') {
      await loadEmailData();
    }
  });

  // Setup view
  on('btn-goto-settings', 'click', async () => {
    showView('settings');
    await loadSettingsView();
  });

  // Form
  on('field-client', 'change', e => { onClientChange(e.target.value); });
  on('btn-create-contact', 'click', handleCreateContact);
  on('field-group',  'change', e => { onGroupChange(e.target.value); });
  on('btn-submit',   'click',  submitTicket);

  // Settings
  on('btn-validate-key',  'click', validateAndSaveApiKey);
  on('btn-save-settings', 'click', saveSettingsFromUI);
  on('btn-back', 'click', async () => {
    if (isAuthenticated()) {
      showView('form');
    } else {
      showView('setup');
    }
  });
  on('btn-show-key', 'click', () => {
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
})();

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
  if (!v) return null;
  const n = parseInt(v, 10);
  return isNaN(n) ? null : n;
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
