function onOpen() {
  const ui = DocumentApp.getUi();

  ui.createMenu('Date Fields')
    .addItem('Insert Field', 'showSidebar')
    .addItem('Update Now', 'updateAllFields')
    .addItem('Toggle Auto Update', 'showSidebar')
    .addToUi();
}

/**
 * SIDEBAR
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Date Field Builder');
  DocumentApp.getUi().showSidebar(html);
}

/**
 * INSERT FIELD
 */
function insertField(config) {
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  if (!cursor) return;

  const id = 'FIELD_' + Date.now();
  const value = renderDate(config);

  cursor.insertText(value);

  PropertiesService.getDocumentProperties()
    .setProperty(id, JSON.stringify({
      mode: config.mode || "MDY",
      lastValue: value
    }));
}

/**
 * TOGGLE AUTO UPDATE (SAFE - NO UI CALLS)
 */
function toggleAutoUpdate() {
  const props = PropertiesService.getDocumentProperties();
  const current = props.getProperty("AUTO_UPDATE_ENABLED") === "true";

  const newState = !current;
  props.setProperty("AUTO_UPDATE_ENABLED", String(newState));

  return newState; // returned to sidebar
}

function isEnabled() {
  return PropertiesService.getDocumentProperties()
    .getProperty("AUTO_UPDATE_ENABLED") === "true";
}

/**
 * UPDATE ENGINE
 */
function updateAllFields() {
  if (!isEnabled()) return;

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const text = body.editAsText();

  const props = PropertiesService.getDocumentProperties().getProperties();

  let content = text.getText();

  Object.keys(props).forEach(id => {
    if (!id.startsWith('FIELD_')) return;

    const data = JSON.parse(props[id]);
    const newValue = renderDate({ mode: data.mode });

    if (data.lastValue && content.includes(data.lastValue)) {
      content = content.split(data.lastValue).join(newValue);

      data.lastValue = newValue;

      PropertiesService.getDocumentProperties()
        .setProperty(id, JSON.stringify(data));
    }
  });

  body.setText(content);
}

/**
 * DATE ENGINE
 */
function renderDate(config) {
  const now = new Date();

  const day = now.getDate();
  const month = now.toLocaleDateString('en-US', { month: 'long' });
  const year = now.getFullYear();

  if (config.mode === "TIME") {
    return "TIME " + now.toLocaleTimeString('en-US');
  }

  if (config.mode === "MY") {
    return month + " " + year;
  }

  if (config.mode === "YM") {
    return year + " " + month;
  }

  if (config.mode === "DMY") {
    return day + " " + month + " " + year;
  }

  return month + " " + day + " " + year;
}
