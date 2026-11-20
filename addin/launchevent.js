Office.onReady();

function onNewMessageComposeHandler(event) {
  // stripes.png is 3000x1 (native 12px stripes).
  // Keep imgWidth = 3000 for exact 12px stripes (or scale proportionally, e.g. 2500 for 10px stripes).
  var imgWidth = 3000;
  var base = 'https://mcmcoproperty.github.io/mcmcopropertyassets/';

  var files = [
    'stripes.png',
    'background.png',
    'sig-header.png',
    'sig-phone.png',
    'sig-web.png',
    'sig-address.png',
    'sig-logo.png'
  ];

  // Default to external URLs; switch each to cid: once attached inline
  var srcMap = {};
  for (var i = 0; i < files.length; i++) {
    srcMap[files[i]] = base + files[i];
  }

  function insertSignature() {
    var deadLinkStyle = 'display:block; border:none; text-decoration:none; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none; line-height:0; font-size:0;';
    var linkRowStyle = 'display:block; font-style:normal; margin:0; padding:0; width:335px; max-width:335px; height:26px; max-height:26px; overflow:hidden; line-height:0; font-size:0; white-space:nowrap;';

    var stripeImg =
      '<a name="sig-stripes" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['stripes.png'] + '" width="' + imgWidth + '" height="40" alt="" draggable="false" style="display:block; border:none; width:' + imgWidth + 'px; min-width:' + imgWidth + 'px; max-width:none; height:40px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a>';

    var leftStackHtml = [
      '<address style="display:block; font-style:normal; margin:0; padding:0; width:335px; max-width:335px; height:15px; max-height:15px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-top" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['background.png'] + '" width="335" height="15" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:15px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
      '<address style="display:block; font-style:normal; margin:0; padding:0; width:335px; max-width:335px; height:69px; max-height:69px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-header" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-header.png'] + '" width="335" height="69" alt="Louise Craighill - Chief Financial Officer" draggable="false" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:69px; font-size:16px; line-height:22px; color:#efdfd4; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
      '<address style="' + linkRowStyle + '"><a href="tel:+61422726186" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;"><img src="' + srcMap['sig-phone.png'] + '" width="335" height="26" alt="+61 422 726 186" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:26px; font-size:12px; line-height:26px; color:#efdfd4;"></a></address>',
      '<address style="' + linkRowStyle + '"><a href="https://mcmanuscoproperty.com.au" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;"><img src="' + srcMap['sig-web.png'] + '" width="335" height="26" alt="mcmanuscoproperty.com.au" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:26px; font-size:12px; line-height:26px; color:#efdfd4;"></a></address>',
      '<address style="' + linkRowStyle + '"><a href="https://maps.google.com/?q=48+Sailors+Bay+Road,+Northbridge+NSW+2063" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;"><img src="' + srcMap['sig-address.png'] + '" width="335" height="26" alt="48 Sailors Bay Road, Northbridge NSW 2063" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:26px; font-size:12px; line-height:26px; color:#efdfd4;"></a></address>',
      '<address style="display:block; font-style:normal; margin:0; padding:0; width:335px; max-width:335px; height:8px; max-height:8px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-bottom" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['background.png'] + '" width="335" height="8" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:8px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>'
    ].join('');

    // Split 170px into 3 strips (57 + 57 + 56) so each stays under Gmail's image-attachment height threshold
    var midStackHtml = [
      '<address style="display:block; font-style:normal; margin:0; padding:0; width:100%; height:57px; max-height:57px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-mid-1" draggable="false" style="width:100%; ' + deadLinkStyle + '"><img src="' + srcMap['background.png'] + '" width="60" height="57" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:100%; min-width:60px; height:57px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
      '<address style="display:block; font-style:normal; margin:0; padding:0; width:100%; height:57px; max-height:57px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-mid-2" draggable="false" style="width:100%; ' + deadLinkStyle + '"><img src="' + srcMap['background.png'] + '" width="60" height="57" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:100%; min-width:60px; height:57px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
      '<address style="display:block; font-style:normal; margin:0; padding:0; width:100%; height:56px; max-height:56px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-mid-3" draggable="false" style="width:100%; ' + deadLinkStyle + '"><img src="' + srcMap['background.png'] + '" width="60" height="56" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:100%; min-width:60px; height:56px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>'
    ].join('');

    var signatureHtml = [
      '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#4f1113" style="width:100%; min-width:670px; table-layout:fixed; border-collapse:separate; border-spacing:0; background-color:#4f1113; color:#efdfd4; font-family:\'Century Gothic\', Helvetica, Arial, sans-serif; -webkit-text-size-adjust:100%;">',
      '  <tr>',
      '    <td width="100%" align="left" bgcolor="#4f1113" style="width:100%; background-color:#4f1113; padding:0; line-height:0; font-size:0;">',
      '      <table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#4f1113" style="width:100%; min-width:670px; table-layout:fixed; border-collapse:separate; border-spacing:0; background-color:#4f1113; line-height:0; font-size:0;">',
      '        <tr>',
      // SECTION 1A: Left 15px padding spacer (15 x 170)
      '          <td width="15" height="170" valign="top" bgcolor="#4f1113" style="width:15px; min-width:15px; max-width:15px; height:170px; padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
      '            <address style="display:block; font-style:normal; margin:0; padding:0; width:15px; max-width:15px; height:170px; max-height:170px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;">',
      '              <a name="sig-bg-left" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['background.png'] + '" width="15" height="170" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:15px; min-width:15px; max-width:15px; height:170px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a>',
      '            </address>',
      '          </td>',
      // SECTION 1B: Fixed Left Info stacked in a single TD (335 x 170 total)
      '          <td width="335" height="170" valign="top" bgcolor="#4f1113" style="width:335px; min-width:335px; max-width:335px; height:170px; padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
                   leftStackHtml,
      '          </td>',
      // SECTION 2: Dynamic Middle Stretch (auto width, min 60px at 670px, stretches to 100%)
      '          <td height="170" valign="top" bgcolor="#4f1113" style="height:170px; padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
                   midStackHtml,
      '          </td>',
      // SECTION 3: Fixed Right Logo + Margin (260 x 170)
      '          <td width="260" height="170" valign="top" bgcolor="#4f1113" style="width:260px; min-width:260px; max-width:260px; height:170px; padding:0; background-color:#4f1113; line-height:0; font-size:0;">',
      '            <address style="display:block; font-style:normal; margin:0; padding:0; width:260px; max-width:260px; height:170px; max-height:170px; overflow:hidden; line-height:0; font-size:0; white-space:nowrap;">',
      '              <a href="https://mcmanuscoproperty.com.au" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;">',
      '                <img src="' + srcMap['sig-logo.png'] + '" width="260" height="170" alt="McManus &amp; Co Property" style="display:block; vertical-align:top; border:none; width:260px; min-width:260px; max-width:260px; height:170px; font-size:14px; line-height:20px; color:#efdfd4;">',
      '              </a>',
      '            </address>',
      '          </td>',
      '        </tr>',
      '      </table>',
      '    </td>',
      '  </tr>',
      '  <tr>',
      '    <td width="100%" height="40" bgcolor="#4a151b" style="width:100%; height:40px; max-height:40px; background-color:#4a151b; padding:0; line-height:0; font-size:0;">',
      '      <address style="display:block; font-style:normal; margin:0; padding:0; width:670px; min-width:100%; height:40px; max-height:40px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;">',
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

  var item = Office.context.mailbox.item;
  if (!item || typeof item.addFileAttachmentAsync !== 'function') {
    insertSignature();
    return;
  }

  var remaining = files.length;

  function attachFile(fileName) {
    item.addFileAttachmentAsync(
      base + fileName,
      fileName,
      { isInline: true },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          srcMap[fileName] = 'cid:' + fileName;
        } else {
          console.error('addFileAttachmentAsync failed for ' + fileName + ':', result.error.message);
        }
        remaining--;
        if (remaining === 0) {
          insertSignature();
        }
      }
    );
  }

  for (var j = 0; j < files.length; j++) {
    attachFile(files[j]);
  }
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
