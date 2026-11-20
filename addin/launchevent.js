Office.onReady();

function onNewMessageComposeHandler(event) {
  // stripes.png is 3000x1 (native 12px stripes).
  // Keep imgWidth = 3000 for exact 12px stripes (or scale proportionally, e.g. 2500 for 10px stripes).
  var imgWidth = 3000;
  var base = 'https://mcmcoproperty.github.io/mcmcopropertyassets/';

  var stripeImg =
    '<img src="' + base + 'stripes.png" width="' + imgWidth + '" height="40" alt="" style="display:block; border:none; width:' + imgWidth + 'px; min-width:' + imgWidth + 'px; max-width:none; height:40px;">';

  var signatureHtml = [
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#4f1113" style="width:100%; min-width:670px; table-layout:fixed; background-color:#4f1113; color:#efdfd4; font-family:\'Century Gothic\', Helvetica, Arial, sans-serif; -webkit-text-size-adjust:100%;">',
    '  <tr>',
    '    <td align="left" bgcolor="#4f1113" style="background-color:#4f1113; padding:0; line-height:0; font-size:0;">',
    '      <table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#4f1113" style="width:100%; min-width:670px; table-layout:fixed; border-collapse:collapse; background-color:#4f1113; line-height:0; font-size:0;">',
    '        <tr>',
    // SECTION 1: Fixed Left Info (350 x 170 total = 15px left + 15px top + 8px bottom padding + 335 x 147 clickable rows)
    '          <td width="350" valign="top" bgcolor="#4f1113" style="width:350px; min-width:350px; max-width:350px; padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
    '            <table cellpadding="0" cellspacing="0" border="0" width="350" bgcolor="#4f1113" style="width:350px; border-collapse:collapse; background-color:#4f1113; line-height:0; font-size:0;">',
    '              <tr>',
    '                <td width="15" height="170" valign="top" style="width:15px; min-width:15px; max-width:15px; padding:0; line-height:0; font-size:0;">',
    '                  <img src="' + base + 'background.png" width="15" height="170" alt="" style="display:block; border:none; width:15px; min-width:15px; max-width:15px; height:170px;">',
    '                </td>',
    '                <td width="335" valign="top" style="width:335px; min-width:335px; max-width:335px; padding:0; line-height:0; font-size:0;">',
    '                  <table cellpadding="0" cellspacing="0" border="0" width="335" bgcolor="#4f1113" style="width:335px; border-collapse:collapse; background-color:#4f1113; line-height:0; font-size:0;">',
    '                    <tr>',
    '                      <td width="335" height="15" style="padding:0; line-height:0; font-size:0;">',
    '                        <img src="' + base + 'background.png" width="335" height="15" alt="" style="display:block; border:none; width:335px; min-width:335px; max-width:335px; height:15px;">',
    '                      </td>',
    '                    </tr>',
    '                    <tr>',
    '                      <td width="335" height="69" style="padding:0; line-height:0; font-size:0;">',
    '                        <img src="' + base + 'sig-header.png" width="335" height="69" alt="Louise Craighill - Chief Financial Officer" style="display:block; border:none; width:335px; min-width:335px; max-width:335px; height:69px;">',
    '                      </td>',
    '                    </tr>',
    '                    <tr>',
    '                      <td width="335" height="26" style="padding:0; line-height:0; font-size:0;">',
    '                        <a href="tel:+61422726186" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;">',
    '                          <img src="' + base + 'sig-phone.png" width="335" height="26" alt="+61 422 726 186" style="display:block; border:none; width:335px; min-width:335px; max-width:335px; height:26px;">',
    '                        </a>',
    '                      </td>',
    '                    </tr>',
    '                    <tr>',
    '                      <td width="335" height="26" style="padding:0; line-height:0; font-size:0;">',
    '                        <a href="https://mcmanuscoproperty.com.au" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;">',
    '                          <img src="' + base + 'sig-web.png" width="335" height="26" alt="mcmanuscoproperty.com.au" style="display:block; border:none; width:335px; min-width:335px; max-width:335px; height:26px;">',
    '                        </a>',
    '                      </td>',
    '                    </tr>',
    '                    <tr>',
    '                      <td width="335" height="26" style="padding:0; line-height:0; font-size:0;">',
    '                        <a href="https://maps.google.com/?q=48+Sailors+Bay+Road,+Northbridge+NSW+2063" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;">',
    '                          <img src="' + base + 'sig-address.png" width="335" height="26" alt="48 Sailors Bay Road, Northbridge NSW 2063" style="display:block; border:none; width:335px; min-width:335px; max-width:335px; height:26px;">',
    '                        </a>',
    '                      </td>',
    '                    </tr>',
    '                    <tr>',
    '                      <td width="335" height="8" style="padding:0; line-height:0; font-size:0;">',
    '                        <img src="' + base + 'background.png" width="335" height="8" alt="" style="display:block; border:none; width:335px; min-width:335px; max-width:335px; height:8px;">',
    '                      </td>',
    '                    </tr>',
    '                  </table>',
    '                </td>',
    '              </tr>',
    '            </table>',
    '          </td>',
    // SECTION 2: Dynamic Middle Stretch (auto width, min 60px at 670px, stretches to 100%)
    '          <td valign="top" bgcolor="#4f1113" style="padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
    '            <img src="' + base + 'background.png" width="60" height="170" alt="" style="display:block; border:none; width:100%; min-width:60px; height:170px;">',
    '          </td>',
    // SECTION 3: Fixed Right Logo + Margin (260 x 170)
    '          <td width="260" valign="top" bgcolor="#4f1113" style="width:260px; min-width:260px; max-width:260px; padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
    '            <a href="https://mcmanuscoproperty.com.au" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;">',
    '              <img src="' + base + 'sig-logo.png" width="260" height="170" alt="McManus &amp; Co Property" style="display:block; border:none; width:260px; min-width:260px; max-width:260px; height:170px;">',
    '            </a>',
    '          </td>',
    '        </tr>',
    '      </table>',
    '    </td>',
    '  </tr>',
    '  <tr>',
    '    <td height="40" bgcolor="#4a151b" style="height:40px; max-height:40px; background-color:#4a151b; padding:0; line-height:0; font-size:0;">',
    '      <address style="display:block; font-style:normal; margin:0; padding:0; width:100%; height:40px; max-height:40px; overflow:hidden; line-height:0; font-size:0;">',
             stripeImg,
    '      </address>',
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
