function doGet(e) {
  // ?manual=1 abre el manual de uso (Manual.html) con la misma dirección del app
  if (e && e.parameter && e.parameter.manual) {
    return HtmlService.createHtmlOutputFromFile('Manual')
      .setTitle('Manual · Comparativo Comercial · Ventus')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Comparativo Comercial · Ventus')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // solo si lo incrustará en Google Sites
}

// Dirección del manual para el botón "Manual" de la barra
function urlManual() {
  const u = ScriptApp.getService().getUrl();
  return u ? u + '?manual=1' : '';
}
