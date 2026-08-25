/**
 * Gorelo Ticket - Settings Manager
 * Persists settings in Office.context.roamingSettings (follows user across devices).
 *
 * Settings stored:
 *   defaultEmailConfigurationId  - Gorelo inbound mailbox ID
 *   defaultTechnicianGroupId     - Default group ID for quick-create
 *   defaultPriorityId            - Default priority (2 = Normal)
 *   defaultSourceId              - Default source (2 = Email)
 *   defaultAssignedToId          - Default assignee (0 = unassigned)
 *   defaultTicketLocation        - Location ID
 *   autoOpenTaskPane             - Open task pane instead of quick-create
 *   lastKnownClientId            - Last used client (for UX convenience)
 */

const DEFAULTS = {
  defaultEmailConfigurationId: 0,
  defaultTechnicianGroupId: 0,
  defaultPriorityId: 2,
  defaultSourceId: 2,
  defaultAssignedToId: 0,
  defaultTicketLocation: 0,
  autoOpenTaskPane: false,
};

export function getSetting(key) {
  if (typeof Office === 'undefined') return DEFAULTS[key];
  const val = Office.context.roamingSettings.get(key);
  return val !== null && val !== undefined ? val : DEFAULTS[key];
}

export function setSetting(key, value) {
  if (typeof Office === 'undefined') return;
  Office.context.roamingSettings.set(key, value);
}

export function saveSettings() {
  return new Promise((resolve, reject) => {
    if (typeof Office === 'undefined') return resolve();
    Office.context.roamingSettings.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

export function getAllSettings() {
  return Object.keys(DEFAULTS).reduce((acc, key) => {
    acc[key] = getSetting(key);
    return acc;
  }, {});
}
