function abrirEscaner() {
  const html = HtmlService
    .createHtmlOutputFromFile("Escaner")
    .setTitle("Escanear compis");

  SpreadsheetApp.getUi().showSidebar(html);
}

