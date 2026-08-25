/**
 * Gorelo Ticket - Task Pane UI
 * Handles the main create-ticket form and settings view.
 */

/* global Office */

import { initAuth, getToken, getCurrentAccount, signOut, isAuthenticated } from '../auth/auth.js';
import {
  getClients, getContacts, getTechnicianGroups, getTechnicians,
  getTags, getAssets, getEmailConfigurations, createTicket,
} from '../gorelo-api.js';
import { getSetting, setSetting, saveSettings, getAllSettings } from '../settings.js';

// ─── State ────────────────────────────────────────────────────────
let allClients = [];
let allContacts = [];
let allGroups = [];
let allTechnicians = [];
let allTags = [];
let allAssets = [];
let allEmailConfigs = [];
let selectedTagIds = [];
let currentEmail = null;

// ─── DOM refs ─────────────────────────────────────────────────────
const $ = id => document.getElementById(id);

// ─── Init ─────────────────────────────────────────────────────────
Office.onReady(async () => {
  try {
    await initAuth();
    wireUpEvents();

    if (isAuthenticated()) {
      await loadEmailData();
      await loadDropdowns();
      updateAccountInfo();
    } else {
      showAuthBanner();
    }
  } catch (err) {
    showStatus('Initialisation error: ' + err.message, 'error');
  }
});

// ─── Event wiring ─────────────────────────────────────────────────
function wireUpEvents() {
  $('btn-settings').addEventListener('click', () => showView('settings'));
  $('btn-back').addEventListener('click', () => showView('create'));
  $('btn-cancel').addEventListener('click', () => Office.context.ui.closeContainer?.());
  $('btn-create').addEventListener('click', handleCreate);
  $('btn-save-settings').addEventListener('click', handleSaveSettings);
  $('btn-login').addEventListener('click', handleLogin);
  $('btn-signout').addEventListener('click', handleSignOut);

  // Client change → filter contacts and assets
  $('field-client').addEventListener('change', handleClientChange);

  // Tag selection → add chip
  $('field-tags').addEventListener('change', handleTagSelect);
}

// ─── View switching ───────────────────────────────────────────────
function showView(name) {
  $('view-create').classList.toggle('hidden', name !== 'create');
  $('view-settings').classList.toggle('hidden', name !== 'settings');
}

// ─── Auth ─────────────────────────────────────────────────────────
function showAuthBanner() {
  $('auth-banner').classList.remove('hidden');
}

async function handleLogin() {
  try {
    showStatus('Signing in to Gorelo…', 'loading');
    await getToken();
    $('auth-banner').classList.add('hidden');
    hideStatus();
    await loadDropdowns();
    updateAccountInfo();
  } catch (err) {
    showStatus('Sign in failed: ' + err.message, 'error');
  }
}

async function handleSignOut() {
  await signOut();
  showAuthBanner();
  updateAccountInfo();
}

function updateAccountInfo() {
  const account = getCurrentAccount();
  $('account-info').textContent = account
    ? `${account.name || account.username} (${account.username})`
    : 'Not signed in';
}

// ─── Load email context ───────────────────────────────────────────
async function loadEmailData() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  currentEmail = {
    subject: item.subject || '',
    from: item.from || { displayName: '', emailAddress: '' },
  };

  $('field-title').value = currentEmail.subject;
  $('from-name').textContent = currentEmail.from.displayName || currentEmail.from.emailAddress;
  $('from-email').textContent = currentEmail.from.displayName ? currentEmail.from.emailAddress : '';

  // Try to auto-match contact by email address
  // (will run after contacts are loaded)
}

// ─── Load dropdowns ───────────────────────────────────────────────
async function loadDropdowns() {
  showStatus('Loading Gorelo data…', 'loading');

  try {
    const [clients, groups, technicians, tags, emailConfigs] = await Promise.allSettled([
      getClients(),
      getTechnicianGroups(),
      getTechnicians(),
      getTags(),
      getEmailConfigurations(),
    ]);

    // Clients
    if (clients.status === 'fulfilled') {
      allClients = normaliseList(clients.value, 'clientId', 'clientName');
      populateSelect('field-client', allClients, 'Select client…');
      populateSelect('s-group', [], 'Loading…'); // reset until client chosen
    }

    // Groups
    if (groups.status === 'fulfilled') {
      allGroups = normaliseList(groups.value, 'technicianGroupId', 'technicianGroupName');
      populateSelect('field-group', allGroups, 'Select group…');
      populateSelect('s-group', allGroups, 'Select group…');
      // Apply saved default
      const savedGroup = getSetting('defaultTechnicianGroupId');
      if (savedGroup) {
        $('field-group').value = savedGroup;
        $('s-group').value = savedGroup;
      }
    }

    // Technicians
    if (technicians.status === 'fulfilled') {
      allTechnicians = normaliseList(technicians.value, 'technicianId', r =>
        `${r.firstName} ${r.lastName}`.trim()
      );
      populateSelect('field-assignee', allTechnicians, 'Unassigned', true);
      populateSelect('s-assignee', allTechnicians, 'Unassigned', true);
      const savedAssignee = getSetting('defaultAssignedToId');
      if (savedAssignee) {
        $('field-assignee').value = savedAssignee;
        $('s-assignee').value = savedAssignee;
      }
    }

    // Tags
    if (tags.status === 'fulfilled') {
      allTags = normaliseList(tags.value, 'tagId', 'tagName');
      populateSelect('field-tags', allTags, 'Add tag…');
    }

    // Email configs
    if (emailConfigs.status === 'fulfilled') {
      allEmailConfigs = normaliseList(emailConfigs.value, 'emailConfigurationId',
        r => r.serviceProviderFromEmailAddress
      );
      populateSelect('field-emailconfig', allEmailConfigs, 'Select mailbox…');
      populateSelect('s-emailconfig', allEmailConfigs, 'Select mailbox…');
      const savedConfig = getSetting('defaultEmailConfigurationId');
      if (savedConfig) {
        $('field-emailconfig').value = savedConfig;
        $('s-emailconfig').value = savedConfig;
      }
    }

    // Apply other saved settings
    const savedPriority = getSetting('defaultPriorityId');
    if (savedPriority) $('field-priority').value = savedPriority;
    $('s-priority').value = getSetting('defaultPriorityId') || 2;
    $('s-autopane').checked = getSetting('autoOpenTaskPane') || false;

    hideStatus();

    // Auto-match sender to contact/client
    if (currentEmail?.from?.emailAddress) {
      await tryAutoMatch(currentEmail.from.emailAddress);
    }

  } catch (err) {
    showStatus('Failed to load data: ' + err.message, 'error');
  }
}

// ─── Client change handler ────────────────────────────────────────
async function handleClientChange() {
  const clientId = parseInt($('field-client').value);
  if (!clientId) {
    populateSelect('field-contact', [], 'Select client first');
    populateSelect('field-asset', [], 'None');
    return;
  }

  try {
    const [contacts, assets] = await Promise.all([
      getContacts(clientId),
      getAssets(clientId),
    ]);

    allContacts = normaliseList(contacts, 'contactId',
      r => `${r.firstName} ${r.lastName}`.trim() || r.emailAddress
    );
    populateSelect('field-contact', allContacts, 'Select contact…', true);

    allAssets = normaliseList(assets, 'assetId', 'assetName');
    populateSelect('field-asset', allAssets, 'None', true);

    // Auto-match contact by email if we have sender info
    if (currentEmail?.from?.emailAddress) {
      tryMatchContact(currentEmail.from.emailAddress);
    }
  } catch (err) {
    console.warn('Failed to load contacts/assets:', err);
  }
}

// ─── Tag handling ─────────────────────────────────────────────────
function handleTagSelect() {
  const select = $('field-tags');
  const id = parseInt(select.value);
  if (!id || selectedTagIds.includes(id)) {
    select.value = '';
    return;
  }
  const tag = allTags.find(t => t.id === id);
  if (!tag) return;

  selectedTagIds.push(id);
  addTagChip(id, tag.name);
  select.value = '';
}

function addTagChip(id, name) {
  const chip = document.createElement('div');
  chip.className = 'tag-chip';
  chip.dataset.tagId = id;
  chip.innerHTML = `<span>${name}</span><button type="button" aria-label="Remove tag">×</button>`;
  chip.querySelector('button').addEventListener('click', () => {
    selectedTagIds = selectedTagIds.filter(t => t !== id);
    chip.remove();
  });
  $('tag-chips').appendChild(chip);
}

// ─── Auto-match sender ────────────────────────────────────────────
async function tryAutoMatch(emailAddress) {
  // Try to find a contact with this email across all clients
  // We'll load all contacts and search
  try {
    const contacts = await getContacts();
    const match = (Array.isArray(contacts) ? contacts : contacts?.data || [])
      .find(c => c.emailAddress?.toLowerCase() === emailAddress.toLowerCase() ||
        (c.secondaryAddress || []).includes(emailAddress.toLowerCase()));

    if (match && match.clientId) {
      $('field-client').value = match.clientId;
      await handleClientChange();
      $('field-contact').value = match.contactId;
    }
  } catch {
    // Silent fail — auto-match is best effort
  }
}

function tryMatchContact(emailAddress) {
  const match = allContacts.find(c => {
    const orig = allContacts._raw?.find(r => r.id === c.id);
    return orig?.emailAddress?.toLowerCase() === emailAddress.toLowerCase();
  });
  if (match) $('field-contact').value = match.id;
}

// ─── Create ticket ────────────────────────────────────────────────
async function handleCreate() {
  const title = $('field-title').value.trim();
  const clientId = parseInt($('field-client').value) || 0;
  const groupId = parseInt($('field-group').value) || 0;

  if (!title) {
    showStatus('Please enter a ticket title.', 'error');
    return;
  }
  if (!groupId) {
    showStatus('Please select a group.', 'error');
    return;
  }

  $('btn-create').disabled = true;
  showStatus('<span class="spinner"></span>Creating ticket…', 'loading');

  try {
    const item = Office.context.mailbox.item;
    const body = await getBodyAsync(item);

    const clientName = allClients.find(c => c.id === clientId)?.name || '';
    const contactId = parseInt($('field-contact').value) || 0;
    const notes = $('field-notes').value.trim();

    const description = buildDescription(
      currentEmail?.from,
      title,
      body,
      notes
    );

    const markdownComment = [
      currentEmail?.from
        ? `**From:** ${currentEmail.from.displayName} <${currentEmail.from.emailAddress}>`
        : '',
      notes ? `**Notes:** ${notes}` : '',
    ].filter(Boolean).join('\n\n');

    await createTicket({
      title,
      description,
      markdownComment,
      clientId,
      clientName,
      contactId,
      outboundContactEmail: currentEmail?.from?.emailAddress || '',
      technicianGroupIds: groupId ? [groupId] : [],
      assignedToId: parseInt($('field-assignee').value) || 0,
      priorityId: parseInt($('field-priority').value) || 2,
      sourceId: parseInt($('field-source').value) || 2,
      emailConfigurationId: parseInt($('field-emailconfig').value) || 0,
      tagIds: selectedTagIds,
      deviceIds: $('field-asset').value ? [$('field-asset').value] : [],
      sendTicketCreatedEmail: $('field-notify').checked,
    });

    showStatus('✓ Ticket created successfully!', 'success');
    $('btn-create').disabled = false;

    // Auto-close after 2 seconds
    setTimeout(() => {
      Office.context.ui.closeContainer?.();
    }, 2000);

  } catch (err) {
    showStatus('Failed to create ticket: ' + err.message, 'error');
    $('btn-create').disabled = false;
  }
}

// ─── Settings ─────────────────────────────────────────────────────
async function handleSaveSettings() {
  setSetting('defaultEmailConfigurationId', parseInt($('s-emailconfig').value) || 0);
  setSetting('defaultTechnicianGroupId', parseInt($('s-group').value) || 0);
  setSetting('defaultAssignedToId', parseInt($('s-assignee').value) || 0);
  setSetting('defaultPriorityId', parseInt($('s-priority').value) || 2);
  setSetting('autoOpenTaskPane', $('s-autopane').checked);

  try {
    await saveSettings();
    showStatus('Settings saved.', 'success');
    setTimeout(hideStatus, 2000);
  } catch (err) {
    showStatus('Failed to save settings: ' + err.message, 'error');
  }
}

// ─── Helpers ──────────────────────────────────────────────────────

/**
 * Normalise an API response into [{id, name}] for populating selects.
 * nameField can be a string key or a function.
 */
function normaliseList(data, idField, nameField) {
  const arr = Array.isArray(data) ? data : (data?.data || data?.items || []);
  return arr.map(r => ({
    id: r[idField],
    name: typeof nameField === 'function' ? nameField(r) : r[nameField],
    _raw: r,
  })).filter(r => r.id);
}

/**
 * Populate a <select> element.
 */
function populateSelect(id, items, placeholder, includeEmpty = false) {
  const select = $(id);
  const current = select.value;
  select.innerHTML = '';

  if (placeholder || includeEmpty) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = placeholder || '—';
    select.appendChild(opt);
  }

  items.forEach(item => {
    const opt = document.createElement('option');
    opt.value = item.id;
    opt.textContent = item.name || `ID: ${item.id}`;
    select.appendChild(opt);
  });

  // Restore selection if still valid
  if (current && [...select.options].some(o => o.value === current)) {
    select.value = current;
  }
}

/**
 * Build HTML ticket description preserving original email context.
 */
function buildDescription(from, subject, body, notes) {
  const fromStr = from
    ? `${escapeHtml(from.displayName)} &lt;${escapeHtml(from.emailAddress)}&gt;`
    : '';

  return `
    <div style="border-left:3px solid #1DAF94;padding:8px 12px;margin-bottom:12px;background:#f0fdf9;">
      <p style="margin:0 0 4px;font-size:12px;color:#6b7280;">Original email</p>
      ${fromStr ? `<p style="margin:0;"><strong>From:</strong> ${fromStr}</p>` : ''}
      ${notes ? `<p style="margin:4px 0 0;"><strong>Notes:</strong> ${escapeHtml(notes)}</p>` : ''}
    </div>
    ${body}
  `.trim();
}

function escapeHtml(str) {
  return (str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function getBodyAsync(item) {
  return new Promise(resolve => {
    item.body.getAsync(Office.CoercionType.Html, {}, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || '');
      } else {
        item.body.getAsync(Office.CoercionType.Text, {}, r => resolve(r.value || ''));
      }
    });
  });
}

function showStatus(html, type) {
  const el = $('status-message');
  el.innerHTML = html;
  el.className = `status-message ${type}`;
  el.classList.remove('hidden');
}

function hideStatus() {
  $('status-message').classList.add('hidden');
}
