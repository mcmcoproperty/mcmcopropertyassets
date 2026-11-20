Office.onReady();

function onNewMessageComposeHandler(event) {
  // stripes.png is 3000x1 (native 12px stripes).
  // Keep imgWidth = 3000 for exact 12px stripes (or scale proportionally, e.g. 2500 for 10px stripes).
  var imgWidth = 3000;
  var base = 'https://mcmcoproperty.github.io/mcmcopropertyassets/';

  // 2px overlap stays >= 1 full device pixel down to 50% scale; set to 1 for tighter spacing.
  var overlap = 2;
  var negMargin = '-' + overlap + 'px';
  var topBgHeight = 15 + overlap;
  var bottomBgHeight = 8 + (overlap * 4);

  var stripeImg =
    '<img src="' + base + 'stripes.png" width="' + imgWidth + '" height="40" alt="" style="display:block; border:none; width:' + imgWidth + 'px; min-width:' + imgWidth + 'px; max-width:none; height:40px;">';

  var leftStackHtml = [
    '<address style="display:block; font-style:normal; margin:0; padding:0; line-height:0; font-size:0;"><img src="' + base + 'background.png" width="335" height="' + topBgHeight + '" alt="" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:' + topBgHeight + 'px;"></address>',
    '<address style="display:block; font-style:normal; margin:' + negMargin + ' 0 0 0; padding:0; line-height:0; font-size:0;"><img src="' + base + 'sig-header.png" width="335" height="69" alt="Louise Craighill - Chief Financial Officer" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:69px;"></address>',
    '<address style="display:block; font-style:normal; margin:' + negMargin + ' 0 0 0; padding:0; line-height:0; font-size:0;"><a href="tel:+61422726186" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;"><img src="' + base + 'sig-phone.png" width="335" height="26" alt="+61 422 726 186" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:26px;"></a></address>',
    '<address style="display:block; font-style:normal; margin:' + negMargin + ' 0 0 0; padding:0; line-height:0; font-size:0;"><a href="https://mcmanuscoproperty.com.au" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;"><img src="' + base + 'sig-web.png" width="335" height="26" alt="mcmanuscoproperty.com.au" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:26px;"></a></address>',
    '<address style="display:block; font-style:normal; margin:' + negMargin + ' 0 0 0; padding:0; line-height:0; font-size:0;"><a href="https://maps.google.com/?q=48+Sailors+Bay+Road,+Northbridge+NSW+2063" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;"><img src="' + base + 'sig-address.png" width="335" height="26" alt="48 Sailors Bay Road, Northbridge NSW 2063" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:26px;"></a></address>',
    '<address style="display:block; font-style:normal; margin:' + negMargin + ' 0 0 0; padding:0; line-height:0; font-size:0;"><img src="' + base + 'background.png" width="335" height="' + bottomBgHeight + '" alt="" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:' + bottomBgHeight + 'px;"></address>'
  ].join('');

  var signatureHtml = [
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#4f1113" style="width:100%; min-width:670px; table-layout:fixed; border-collapse:separate; border-spacing:0; background-color:#4f1113; color:#efdfd4; font-family:\'Century Gothic\', Helvetica, Arial, sans-serif; -webkit-text-size-adjust:100%;">',
    '  <tr>',
    '    <td align="left" bgcolor="#4f1113" style="background-color:#4f1113; padding:0; line-height:0; font-size:0;">',
    '      <table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#4f1113" style="width:100%; min-width:670px; table-layout:fixed; border-collapse:separate; border-spacing:0; background-color:#4f1113; line-height:0; font-size:0;">',
    '        <tr>',
    // SECTION 1A: Left 15px padding spacer (overlaps 2px right under Section 1B)
    '          <td width="15" height="170" valign="top" bgcolor="#4f1113" style="width:15px; min-width:15px; max-width:15px; height:170px; padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
    '            <address style="display:block; font-style:normal; margin:0 ' + negMargin + ' 0 0; padding:0; height:170px; line-height:0; font-size:0;">',
    '              <img src="' + base + 'background.png" width="15" height="170" alt="" style="display:block; vertical-align:top; border:none; width:100%; height:170px;">',
    '            </address>',
    '          </td>',
    // SECTION 1B: Fixed Left Info stacked in a single TD (335 x 170 total)
    '          <td width="335" height="170" valign="top" bgcolor="#4f1113" style="width:335px; min-width:335px; max-width:335px; height:170px; padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
                 leftStackHtml,
    '          </td>',
    // SECTION 2: Dynamic Middle Stretch (overlaps 2px left over Col 1B right margin, 2px right under Logo)
    '          <td height="170" valign="top" bgcolor="#4f1113" style="height:170px; padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
    '            <address style="display:block; font-style:normal; margin:0 ' + negMargin + '; padding:0; height:170px; line-height:0; font-size:0;">',
    '              <img src="' + base + 'background.png" width="60" height="170" alt="" style="display:block; vertical-align:top; border:none; width:100%; min-width:60px; height:170px;">',
    '            </address>',
    '          </td>',
    // SECTION 3: Fixed Right Logo + Margin (260 x 170)
    '          <td width="260" height="170" valign="top" bgcolor="#4f1113" style="width:260px; min-width:260px; max-width:260px; height:170px; padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
    '            <a href="https://mcmanuscoproperty.com.au" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;">',
    '              <img src="' + base + 'sig-logo.png" width="260" height="170" alt="McManus &amp; Co Property" style="display:block; vertical-align:top; border:none; width:260px; min-width:260px; max-width:260px; height:170px;">',
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
