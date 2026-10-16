/* settings.js — persists configuration in Office RoamingSettings so it
   follows the user across devices (desktop, web, Mac, mobile). */

const DEFAULTS = {
  apiKey: '',
  defaultTechnicianGroupId: '',
  defaultPriorityId: 2,          // Normal
  defaultSourceId: 2,            // Email
  defaultAssignedToId: '',
  defaultTicketLocation: 0,
  defaultEmailConfigurationId: '',
  autoOpenTaskPane: false,
};

function getSettings() {
  try {
    const rs = Office.context.roamingSettings;
    const stored = rs.get('gorelo_settings');
    return stored ? { ...DEFAULTS, ...stored } : { ...DEFAULTS };
  } catch {
    return { ...DEFAULTS };
  }
}

export function getSetting(key) {
  return getSettings()[key];
}

export function saveSettings(updates) {
  return new Promise((resolve, reject) => {
    try {
      const rs = Office.context.roamingSettings;
      const current = getSettings();
      const merged = { ...current, ...updates };
      rs.set('gorelo_settings', merged);
      rs.saveAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(merged);
        } else {
          reject(new Error('Failed to save settings: ' + result.error.message));
        }
      });
    } catch (err) {
      reject(err);
    }
  });
}

export function getAllSettings() {
  return getSettings();
}
