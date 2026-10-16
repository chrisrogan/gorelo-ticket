/* gorelo-api.js — All calls to the Gorelo Public API (api.usw.gorelo.io).
   Authentication: X-API-Key header using the key stored in RoamingSettings. */

import { getApiKey } from './auth/auth.js';

const BASE = 'https://api.usw.gorelo.io';

function headers() {
  return {
    'X-API-Key': getApiKey(),
    'Content-Type': 'application/json',
    'Accept': 'application/json',
  };
}

async function apiGet(path) {
  const res = await fetch(BASE + path, { headers: headers() });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Gorelo API ${res.status}: ${text || res.statusText}`);
  }
  return res.json();
}

async function apiPost(path, body) {
  const res = await fetch(BASE + path, {
    method: 'POST',
    headers: headers(),
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Gorelo API ${res.status}: ${text || res.statusText}`);
  }
  // Some endpoints return 204 No Content
  const ct = res.headers.get('content-type') || '';
  if (res.status === 204 || !ct.includes('application/json')) return null;
  return res.json();
}

// ─── Lookups ─────────────────────────────────────────────────────────────────

export async function getClients(search = '') {
  const q = search ? `&Search=${encodeURIComponent(search)}` : '';
  const data = await apiGet(`/v1/clients?PageSize=50&PageNumber=1${q}`);
  // API returns { items: [...] } or a plain array depending on version
  return Array.isArray(data) ? data : (data.items ?? data.data ?? []);
}

export async function getContacts(clientId) {
  const q = clientId ? `?ClientId=${clientId}&PageSize=100&PageNumber=1` : '?PageSize=100&PageNumber=1';
  const data = await apiGet(`/v1/contacts${q}`);
  return Array.isArray(data) ? data : (data.items ?? data.data ?? []);
}

export async function getTechnicianGroups() {
  const data = await apiGet('/v1/technician-groups?PageSize=100&PageNumber=1');
  return Array.isArray(data) ? data : (data.items ?? data.data ?? []);
}

export async function getTechnicians(groupId) {
  const q = groupId ? `?TechnicianGroupId=${groupId}&PageSize=100&PageNumber=1` : '?PageSize=100&PageNumber=1';
  const data = await apiGet(`/v1/technicians${q}`);
  return Array.isArray(data) ? data : (data.items ?? data.data ?? []);
}

export async function getTags() {
  const data = await apiGet('/v1/tags?PageSize=200&PageNumber=1');
  return Array.isArray(data) ? data : (data.items ?? data.data ?? []);
}

export async function getEmailConfigurations() {
  const data = await apiGet('/v1/email-configurations?PageSize=50&PageNumber=1');
  return Array.isArray(data) ? data : (data.items ?? data.data ?? []);
}

// Static lookups — these don't change and match the values confirmed from
// the Gorelo Vuex store / network capture.
export const PRIORITIES = [
  { id: 1, name: 'High' },
  { id: 2, name: 'Normal' },
  { id: 3, name: 'Low' },
  { id: 4, name: 'Urgent' },
  { id: 5, name: 'No Priority' },
];

export const SOURCES = [
  { id: 1, name: 'Web' },
  { id: 2, name: 'Email' },
  { id: 3, name: 'Phone' },
  { id: 4, name: 'Chat' },
  { id: 5, name: 'Alert' },
  { id: 6, name: 'API' },
];

// ─── Ticket creation ──────────────────────────────────────────────────────────

/**
 * Creates a ticket via POST /v1/tickets.
 *
 * @param {object} opts
 * @param {string}   opts.title
 * @param {string}   opts.description        Plain text / HTML body
 * @param {string}   opts.markdownComment    Markdown version of body (can equal description)
 * @param {number}   [opts.clientId]
 * @param {string}   [opts.clientName]
 * @param {number}   [opts.contactId]
 * @param {string}   [opts.outboundContactEmail]
 * @param {number[]} [opts.technicianGroupIds]
 * @param {number}   [opts.assignedToId]
 * @param {number}   [opts.priorityId]       1=High 2=Normal 3=Low 4=Urgent
 * @param {number}   [opts.sourceId]         2=Email
 * @param {string[]} [opts.newTags]          Tag names to create
 * @param {string[]} [opts.existingTagIds]   Existing tag IDs
 * @param {number}   [opts.emailConfigurationId]
 * @param {boolean}  [opts.sendNotification]
 */
export async function createTicket(opts) {
  const payload = {
    TicketTitle: opts.title,
    TicketDescription: opts.description || '',
    markdownComment: opts.markdownComment || opts.description || '',
    ClientId: opts.clientId || 0,
    ClientName: opts.clientName || '',
    ContactId: opts.contactId || 0,
    outboundContactEmail: opts.outboundContactEmail || '',
    technicianGroupIds: opts.technicianGroupIds || [],
    AssignedToId: opts.assignedToId || 0,
    TicketPriorityId: opts.priorityId || 2,
    TicketSeverityId: null,
    TicketSourceId: opts.sourceId || 2,      // Email
    TicketStatusId: 1,                        // Open
    emailConfigurationId: opts.emailConfigurationId || null,
    ticketLocation: opts.ticketLocation || 0,
    NewTicketTags: opts.newTags || [],
    OldTicketTags: opts.existingTagIds || [],
    TicketDeviceIds: [],
    sendTicketCreatedEmail: opts.sendNotification !== false,
    isFromWeb: true,
  };

  return apiPost('/v1/tickets', payload);
}
