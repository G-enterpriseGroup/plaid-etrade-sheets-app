function showPlaidConnectModal() {
  const template = HtmlService.createTemplateFromFile('PlaidLinkModal');
  template.linkToken = createPlaidLinkToken();
  template.mode = 'create';
  template.itemId = '';

  const html = template.evaluate()
    .setWidth(720)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Connect Brokerage');
  return { status: 'opened' };
}

function showPlaidUpdateModal(itemId) {
  const template = HtmlService.createTemplateFromFile('PlaidLinkModal');
  template.linkToken = createPlaidUpdateLinkToken(itemId);
  template.mode = 'update';
  template.itemId = itemId || '';

  const html = template.evaluate()
    .setWidth(720)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Update Brokerage Login');
  return { status: 'opened' };
}
