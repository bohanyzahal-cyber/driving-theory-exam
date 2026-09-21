// ========== Office WhatsApp number (sender) ==========
// Dedicated number for outbound WhatsApp messages to examinees, registered
// 2026-05-15. Currently inactive — used only as a future-ready config for
// when the Meta Cloud API is wired up. Read via getOfficeWhatsAppNumber()
// so it can be overridden at runtime through Script Properties without redeploy.
var OFFICE_WHATSAPP_NUMBER_DEFAULT = '0529151157';
function getOfficeWhatsAppNumber() {
  try {
    var prop = PropertiesService.getScriptProperties().getProperty('OFFICE_WHATSAPP_NUMBER');
    if (prop && String(prop).trim()) return String(prop).trim();
  } catch (e) {}
  return OFFICE_WHATSAPP_NUMBER_DEFAULT;
}

