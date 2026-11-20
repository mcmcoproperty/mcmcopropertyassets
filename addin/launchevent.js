Office.onReady();

function onNewMessageComposeHandler(event) {
  // Adjust stripe size while keeping total height = 40px:
  // Narrower stripes: thickness = 4, rows = 10
  // Medium stripes:   thickness = 5, rows = 8
  // Wider stripes:    thickness = 8, rows = 5
  var stripeThickness = 5;
  var stripeRows = 8;

  var stripeRow =
    '<tr>' +
      '<td bgcolor="#4a151b" style="height:0; padding:0; font-size:0; line-height:0; background-color:#4a151b; border-top:' + stripeThickness + 'px dashed #efdfd4; border-right:none; border-bottom:none; border-left:none;">' +
        '<div style="height:0; font-size:0; line-height:0;">&nbsp;</div>' +
      '</td>' +
    '</tr>';

  var stripeRowsHtml = '';
  for (var i = 0; i < stripeRows; i++) {
    stripeRowsHtml += stripeRow;
  }

  var signatureHtml = [
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#4f1113" style="width:100%; min-width:670px; table-layout:fixed; background-color:#4f1113; color:#efdfd4; font-family:\'Century Gothic\', Helvetica, Arial, sans-serif; -webkit-text-size-adjust:100%;">',
    '  <tr>',
    '    <td align="left" bgcolor="#4f1113" style="background-color:#4f1113; padding:15px 60px 15px 30px;">',
    '      <table cellpadding="0" cellspacing="0" border="0" width="100%" style="width:100%; min-width:580px; table-layout:auto;">',
    '        <tr>',
    '          <td width="320" align="left" valign="middle" style="width:320px; min-width:320px;">',
    '            <table cellpadding="0" cellspacing="0" border="0" style="color:#efdfd4; font-family:\'Century Gothic\', Helvetica, Arial, sans-serif;">',
    '              <tr><td style="font-size:22px; font-weight:normal; letter-spacing:0.5px; white-space:nowrap;">Louise&nbsp;Craighill</td></tr>',
    '              <tr><td style="font-size:13px; font-style:italic; padding-bottom:25px; white-space:nowrap;">Chief&nbsp;Financial&nbsp;Officer</td></tr>',
    '              <tr><td>',
    '                <table cellpadding="0" cellspacing="0" border="0" style="font-size:12px; font-family:\'Century Gothic\', Helvetica, Arial, sans-serif;">',
    '                  <tr>',
    '                    <td width="18" valign="middle" style="width:18px; min-width:18px; max-width:18px; padding-bottom:8px;"><img src="https://mcmcoproperty.github.io/mcmcopropertyassets/phone.png" width="18" height="18" alt="Phone" style="display:block; width:18px; height:18px;"></td>',
    '                    <td width="12" style="width:12px; min-width:12px; padding-bottom:8px;">&nbsp;</td>',
    '                    <td valign="middle" style="padding-bottom:8px; white-space:nowrap;"><a href="tel:+61422726186" style="color:#efdfd4; text-decoration:none;">+61&nbsp;422&nbsp;726&nbsp;186</a></td>',
    '                  </tr>',
    '                </table>',
    '                <table cellpadding="0" cellspacing="0" border="0" style="font-size:12px; font-family:\'Century Gothic\', Helvetica, Arial, sans-serif;">',
    '                  <tr>',
    '                    <td width="18" valign="middle" style="width:18px; min-width:18px; max-width:18px; padding-bottom:8px;"><img src="https://mcmcoproperty.github.io/mcmcopropertyassets/website.png" width="18" height="18" alt="Web" style="display:block; width:18px; height:18px;"></td>',
    '                    <td width="12" style="width:12px; min-width:12px; padding-bottom:8px;">&nbsp;</td>',
    '                    <td valign="middle" style="padding-bottom:8px; white-space:nowrap;"><a href="https://mcmanuscoproperty.com.au" style="color:#efdfd4; text-decoration:none;">mcmanuscoproperty.com.au&zwnj;</a></td>',
    '                  </tr>',
    '                </table>',
    '                <table cellpadding="0" cellspacing="0" border="0" style="font-size:12px; font-family:\'Century Gothic\', Helvetica, Arial, sans-serif;">',
    '                  <tr>',
    '                    <td width="18" valign="top" style="width:18px; min-width:18px; max-width:18px; padding-top:2px;"><img src="https://mcmcoproperty.github.io/mcmcopropertyassets/location.png" width="18" height="18" alt="Location" style="display:block; width:18px; height:18px;"></td>',
    '                    <td width="12" style="width:12px; min-width:12px; padding-top:2px;">&nbsp;</td>',
    '                    <td valign="middle" style="line-height:16px; white-space:nowrap;"><a href="https://maps.google.com/?q=48+Sailors+Bay+Road,+Northbridge+NSW+2063" style="color:#efdfd4; text-decoration:none;">&zwnj;48 Sailors Bay Road, Northbridge NSW 2063</a></td>',
    '                  </tr>',
    '                </table>',
    '              </td></tr>',
    '            </table>',
    '          </td>',
    '          <td align="right" valign="middle" style="padding-left:40px;">',
    '            <a href="https://mcmanuscoproperty.com.au" style="display:inline-block;"><img src="https://mcmcoproperty.github.io/mcmcopropertyassets/brandmark.png" width="200" alt="McManus &amp; Co Property" style="display:block; border:none; width:200px; min-width:200px; max-width:200px; height:auto; margin-left:auto;"></a>',
    '          </td>',
    '        </tr>',
    '      </table>',
    '    </td>',
    '  </tr>',
    '  <tr>',
    '    <td height="40" bgcolor="#4a151b" style="height:40px; background-color:#4a151b; padding:0; line-height:0; font-size:0;">',
    '      <table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#4a151b" style="width:100%; border-collapse:separate; background-color:#4a151b; line-height:0; font-size:0;">',
             stripeRowsHtml,
    '      </table>',
    '    </td>',
    '  </tr>',
    '</table>'
  ].join('');

  Office.context.mailbox.item.body.setSignatureAsync(
    signatureHtml,
    { coercionType: Office.CoercionType.Html },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error('setSignatureAsync failed:', result.error.message);
      }
      event.completed();
    }
  );
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
