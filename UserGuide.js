// Server-side helper to show the user guide dialog
function showUserGuide() {
  try {
    // Load the markdown content from the separate file
    const markdown = HtmlService.createHtmlOutputFromFile('UserGuideContent').getContent();
    
    // Create the template and pass the markdown content
    const template = HtmlService.createTemplateFromFile('UserGuideDialog');
    template.markdown = JSON.stringify(markdown); // Stringify to safely pass as JS variable
    
    const html = template.evaluate()
      .setWidth(900)
      .setHeight(700);
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Gebruikershandleiding ThermoClinics');
  } catch (err) {
    Logger.log('Error showing User Guide: ' + err.toString());
    SpreadsheetApp.getUi().alert('Fout bij het openen van de handleiding: ' + err.message);
  }
}
