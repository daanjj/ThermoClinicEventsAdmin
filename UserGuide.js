// Server-side helper to show the user guide dialog
/**
 * Include helper for HTML templating. Returns the raw content of a project file.
 * Tries the given name first, then with a .html suffix, then with the exact path.
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    try {
      return HtmlService.createHtmlOutputFromFile(filename + '.html').getContent();
    } catch (e2) {
      // Last resort: return an empty string to avoid template errors
      return '';
    }
  }
}

function showUserGuide() {
  const template = HtmlService.createTemplateFromFile('UserGuideDialog');
  const html = template.evaluate().setWidth(900).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gebruikershandleiding ThermoClinics');
}
