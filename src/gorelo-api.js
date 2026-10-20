/* gorelo-api.js — All calls to the Gorelo Public API (api.usw.gorelo.io).
   Authentication: X-API-Key header using the key stored in RoamingSettings.

   Real API shape (confirmed from live responses):
   - Response wrapper: { StatusCode, IsSuccess, Data: [...], DataContext: { Pagination: { NextCursor, HasMore } } }
   - Pagination: cursor-based using ?Cursor=<token>&PageSize=<n>  (no PageNumber)
   - Clients:   Data[].{ Id, Name, Status, Domains[] }
   - Contacts:  Data[].{ Id, FirstName, LastName, ClientId, PrimaryEmail }
   - Users:     Data[].{ Id, FirstName, LastName, Email }
   - Groups:    /v1/organization/groups  (no pagination params)
   - Tags:      /v1/tickets/tags  (requires Tickets read scope)
   - Tickets:   POST /v1/tickets
*/

const _v = new URL(import.meta.url).searchParams.get('v') || Date.now();
const { getApiKey } = await import(`./auth/auth.js?v=${_v}`);

const BASE = 'https://gorelo-proxy.chris-rogan.workers.dev';

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
  const json = await res.json();
  if (!json.IsSuccess) {
    const msg = json.Notifications?.[0]?.Message || `Status ${json.StatusCode}`;
    throw new Error(`Gorelo: ${msg}`);
  }
  return json;
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
  const ct = res.headers.get('content-type') || '';
  if (res.status === 204 || !ct.includes('application/json')) return null;
  const json = await res.json();
  if (json.IsSuccess === false) {
    const msg = json.Notifications?.[0]?.Message || `Status ${json.StatusCode}`;
    throw new Error(`Gorelo: ${msg}`);
  }
  return json.Data ?? json;
}

/** Unwrap Data array from standard response */
function unwrap(json) {
  const data = json.Data ?? json.data ?? json;
  const arr = Array.isArray(data) ? data : [];
  if (arr.length > 0) {
    console.log('[Gorelo] sample item keys:', Object.keys(arr[0]));
  }
  return arr;
}

// ─── Lookups ─────────────────────────────────────────────────────────────────

/** Read a field that may be PascalCase or camelCase */
function f(obj, ...keys) {
  for (const k of keys) {
    if (obj[k] !== undefined && obj[k] !== null && obj[k] !== '') return obj[k];
  }
  return '';
}

export async function getClients(search = '') {
  const q = search ? `&Search=${encodeURIComponent(search)}` : '';
  const json = await apiGet(`/v1/clients?PageSize=200${q}`);
  return unwrap(json).map(c => ({
    id: f(c, 'Id', 'id', 'clientId', 'ClientId'),
    name: f(c, 'Name', 'name', 'clientName', 'ClientName'),
  }));
}

export async function getContacts(clientId) {
  const q = clientId ? `?ClientId=${clientId}&PageSize=200` : '?PageSize=200';
  const json = await apiGet(`/v1/contacts${q}`);
  return unwrap(json).map(c => {
    const first = f(c, 'FirstName', 'firstName');
    const last = f(c, 'LastName', 'lastName');
    const email = f(c, 'PrimaryEmail', 'primaryEmail', 'emailAddress', 'EmailAddress');
    return {
      id: f(c, 'Id', 'id', 'contactId', 'ContactId'),
      name: [first, last].filter(Boolean).join(' ') || email || `Contact #${c.Id || c.id}`,
      clientId: f(c, 'ClientId', 'clientId'),
      email,
    };
  });
}

export async function getTechnicianGroups() {
  const json = await apiGet('/v1/organization/groups');
  return unwrap(json).map(g => ({
    id: f(g, 'Id', 'id', 'groupId', 'GroupId'),
    name: f(g, 'Name', 'name'),
  }));
}

export async function getTechnicians(groupId) {
  const q = groupId ? `?GroupId=${groupId}&PageSize=200` : '?PageSize=200';
  try {
    const json = await apiGet(`/v1/organization/users${q}`);
    return unwrap(json).map(u => {
      const first = f(u, 'FirstName', 'firstName');
      const last = f(u, 'LastName', 'lastName');
      const email = f(u, 'Email', 'email');
      return {
        id: f(u, 'Id', 'id', 'userId', 'UserId'),
        name: [first, last].filter(Boolean).join(' ') || email || `User #${u.Id || u.id}`,
        email,
      };
    });
  } catch {
    const json = await apiGet('/v1/organization/users?PageSize=200');
    return unwrap(json).map(u => {
      const first = f(u, 'FirstName', 'firstName');
      const last = f(u, 'LastName', 'lastName');
      const email = f(u, 'Email', 'email');
      return {
        id: f(u, 'Id', 'id', 'userId', 'UserId'),
        name: [first, last].filter(Boolean).join(' ') || email || `User #${u.Id || u.id}`,
        email,
      };
    });
  }
}

export async function getTags() {
  // Requires 'Tickets' read scope on the API key.
  // Returns empty array gracefully if the key lacks that permission.
  try {
    const json = await apiGet('/v1/tickets/tags?PageSize=200');
    return unwrap(json).map(t => ({ id: t.Id, name: t.Name || t.TagName || '' }));
  } catch {
    return [];
  }
}

export async function getEmailConfigurations() {
  // Not present in the Public API — return empty array.
  return [];
}

// Static lookups — values confirmed from Gorelo UI
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
 * @param {string}   opts.markdownComment    Markdown version of body
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
  // Strip problematic Unicode (zero-width spaces, soft hyphens, BOM, C0 controls)
  // that cause Gorelo's API to return 400 "Invalid or malformed request body"
  function sanitise(str) {
    if (!str) return '';
    return str
      .replace(/<[^>]*>/g, ' ')                                              // strip HTML tags
      .replace(/&[a-z]+;|&#[0-9]+;|&#x[0-9a-f]+;/gi, ' ')                  // HTML entities
      .replace(/[\u00ad\u200b-\u200f\u2028\u2029\ufeff\u180e]/g, '') // soft-hyphen, ZWS
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, '') // C0 controls
      .replace(/\s+/g, ' ')                                                  // collapse whitespace
      .trim();
  }

  // Build payload omitting null/empty optional fields to avoid Gorelo 400s
  const payload = {
    TicketTitle: sanitise(opts.title),
    TicketDescription: sanitise(opts.description),
    TicketPriorityId: opts.priorityId || 2,
    TicketSourceId: opts.sourceId || 2,
    TicketStatusId: 1,
    SendTicketCreatedEmail: opts.sendNotification !== false,
  };

  // Only include optional fields when they have values
  if (opts.clientId)                   payload.ClientId = opts.clientId;
  if (opts.contactId)                  payload.ContactId = opts.contactId;
  if (opts.technicianGroupIds?.length) payload.TechnicianGroupId = opts.technicianGroupIds[0];
  if (opts.assignedToId)               payload.AssignedToId = opts.assignedToId;
  if (opts.emailConfigurationId)       payload.EmailConfigurationId = opts.emailConfigurationId;
  if (opts.newTags?.length)            payload.NewTicketTags = opts.newTags;
  if (opts.existingTagIds?.length)     payload.OldTicketTags = opts.existingTagIds;

  // Truncate description to avoid hitting Gorelo's field length limit
  if (payload.TicketDescription && payload.TicketDescription.length > 2000) {
    payload.TicketDescription = payload.TicketDescription.substring(0, 1997) + '...';
  }

  console.log('[Gorelo] createTicket payload:', JSON.stringify(payload));

  // DIAGNOSTIC: try bare-minimum payload first to isolate whether 400 is
  // caused by field names/values or by description content
  const minPayload = {
    TicketTitle: payload.TicketTitle,
    TicketDescription: payload.TicketDescription,
    TicketPriorityId: payload.TicketPriorityId,
    TicketSourceId: payload.TicketSourceId,
    TicketStatusId: payload.TicketStatusId,
    SendTicketCreatedEmail: payload.SendTicketCreatedEmail,
  };
  console.log('[Gorelo] trying minimal payload:', JSON.stringify(minPayload));
  try {
    const result = await apiPost('/v1/tickets', minPayload);
    console.log('[Gorelo] minimal payload SUCCEEDED');
    return result;
  } catch (minErr) {
    console.warn('[Gorelo] minimal payload ALSO failed:', minErr.message);
  }

  console.log('[Gorelo] trying full payload...');
  return apiPost('/v1/tickets', payload);
}
