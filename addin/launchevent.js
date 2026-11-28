Office.onReady();

// Per-employee directory keyed by lowercase Microsoft 365 email address.
// Shared assets (stripes, bg-left, bg-mid, web, address, logo) stay in the root folder.
// For future staff, set folder to e.g. 'staff/james/' and upload their sig-header.png and sig-phone.png there.
var STAFF = {
  'louise@mcmanuscoproperty.com': {
    folder: 'staff/louise/',
    headerAlt: 'Louise Craighill - Chief Financial Officer',
    phoneHref: 'tel:+61422726186',
    phoneAlt: '+61 422 726 186'
  }
    /*
  'john@mcmanuscoproperty.com': {
    folder: 'staff/john/',
    headerAlt: 'John McManus - Director',
    phoneHref: 'tel:+61425231131',
    phoneAlt: '+61 425 231 131'
  },
  'morgan@mcmanuscoproperty.com': {
    folder: 'staff/morgan/',
    headerAlt: 'Morgan McManus - Co-Founder',
    phoneHref: 'tel:+61424402806',
    phoneAlt: '+61 424 402 806'
  },
  'matthew@mcmanuscoproperty.com': {
    folder: 'staff/matthew/',
    headerAlt: 'Matthew Downes',
    phoneHref: 'tel:+61497770072',
    phoneAlt: '+61 497 770 072'
  },
  'rebecca@mcmanuscoproperty.com': {
    folder: 'staff/rebecca/',
    headerAlt: 'Bec Blair',
    phoneHref: 'tel:+61409634785',
    phoneAlt: '+61 409 634 785'
  },
  'info@mcmanuscoproperty.com': {
    folder: 'staff/info/',
    headerAlt: 'Reception',
    phoneHref: 'tel:+61288536255',
    phoneAlt: '+61 2 8853 6255'
  },
  'accounts@mcmanuscoproperty.com': {
    folder: 'staff/accounts/',
    headerAlt: 'Accounts',
    phoneHref: 'tel:+61288536255',
    phoneAlt: '+61 2 8853 6255'
  },
  'trust@mcmanuscoproperty.com': {
    folder: 'staff/accounts/',
    headerAlt: 'Accounts',
    phoneHref: 'tel:+61288536255',
    phoneAlt: '+61 2 8853 6255'
  }
  */
};

var SIG_FILE_NAMES = {
  'sig-stripes.png': true,
  'sig-background.png': true,
  'sig-bg-pad.png': true,
  'sig-bg-left.png': true,
  'sig-bg-mid.png': true,
  'sig-header.png': true,
  'sig-phone.png': true,
  'sig-web.png': true,
  'sig-address.png': true,
  'sig-logo.png': true
};

function arrayBufferToBase64(buffer) {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  var bytes = new Uint8Array(buffer);
  var len = bytes.length;
  var out = '';
  for (var i = 0; i < len; i += 3) {
    var b1 = bytes[i];
    var b2 = i + 1 < len ? bytes[i + 1] : 0;
    var b3 = i + 2 < len ? bytes[i + 2] : 0;
    out += chars[b1 >> 2];
    out += chars[((b1 & 3) << 4) | (b2 >> 4)];
    out += i + 1 < len ? chars[((b2 & 15) << 2) | (b3 >> 6)] : '=';
    out += i + 2 < len ? chars[b3 & 63] : '=';
  }
  return out;
}

function extractTextStyleFromCss(cssText, fallbackSize) {
  if (!cssText) {
    return null;
  }

  var famMatch = cssText.match(/font-family\s*:\s*([^;]+)/i);
  if (!famMatch) {
    return null;
  }

  var fontFamily = famMatch[1].trim().replace(/"/g, "'").replace(/[<>]/g, '');
  if (!fontFamily || /century gothic/i.test(fontFamily)) {
    return null;
  }

  var sizeMatch = cssText.match(/font-size\s*:\s*([^;]+)/i);
  var fontSize = sizeMatch ? sizeMatch[1].trim().replace(/["'<>]/g, '') : (fallbackSize || '');
  if (/^0+(?:\.0+)?(?:px|pt|em|rem|%)?$/i.test(fontSize)) {
    fontSize = fallbackSize || '';
  }

  var colorMatch = cssText.match(/(?:^|[;\s])color\s*:\s*([^;]+)/i);
  var color = colorMatch ? colorMatch[1].trim().replace(/["'<>]/g, '') : '';
  if (!color || /^(windowtext|inherit|initial|auto)$/i.test(color)) {
    color = '#000000';
  }

  var out = 'font-family:' + fontFamily + ';';
  if (fontSize) {
    out += ' font-size:' + fontSize + ';';
  }
  out += ' color:' + color + ';';
  return out;
}

function detectDefaultTextStyle(item, callback) {
  var fallback = 'font-family:Aptos, Calibri, Helvetica, Arial, sans-serif; font-size:12pt; color:#000000;';

  if (!item || !item.body || typeof item.body.getAsync !== 'function') {
    callback(fallback);
    return;
  }

  item.body.getAsync(Office.CoercionType.Html, function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
      callback(fallback);
      return;
    }

    var html = result.value;

    // Inspect only the user's compose area above any existing signature or quoted thread
    var bodyStart = html.search(/<body\b[^>]*>/i);
    var bodyContent = bodyStart !== -1 ? html.slice(bodyStart) : html;
    var cutIdx = bodyContent.search(/id=["']?(?:x_)?(?:Signature|appendonsend|divRplyFwdMsg)["']?|<blockquote\b/i);
    if (cutIdx !== -1) {
      bodyContent = bodyContent.slice(0, cutIdx);
    }

    // 1. Outlook Web / New Outlook / Mac: check the initial elementToProof compose wrapper
    var etpMatch =
      bodyContent.match(/<[^>]*elementToProof[^>]*style="([^"]*)"/i) ||
      bodyContent.match(/<[^>]*style="([^"]*)"[^>]*elementToProof/i);
    if (etpMatch) {
      var etpStyle = extractTextStyleFromCss(etpMatch[1]);
      if (etpStyle) {
        callback(etpStyle);
        return;
      }
    }

    // 2. Fallback scan of inline style attributes inside the compose area
    var styleRegex = /style="([^"]*)"/gi;
    var m;
    while ((m = styleRegex.exec(bodyContent)) !== null) {
      var inlineStyle = extractTextStyleFromCss(m[1]);
      if (inlineStyle) {
        callback(inlineStyle);
        return;
      }
    }

    // 3. Classic Outlook (Windows): read Word HTML stylesheet rules (EmailStyle / MsoNormal)
    var msoMatch = html.match(/MsoNormal[^{]*\{([^}]*)\}/i);
    var msoSizeMatch = msoMatch ? msoMatch[1].match(/font-size\s*:\s*([^;]+)/i) : null;
    var msoSize = msoSizeMatch ? msoSizeMatch[1].trim().replace(/["'<>]/g, '') : '';

    var emailStyleMatch = html.match(/EmailStyle\d+[^{]*\{([^}]*)\}/i);
    if (emailStyleMatch) {
      var emailCssStyle = extractTextStyleFromCss(emailStyleMatch[1], msoSize);
      if (emailCssStyle) {
        callback(emailCssStyle);
        return;
      }
    }

    if (msoMatch) {
      var msoStyle = extractTextStyleFromCss(msoMatch[1]);
      if (msoStyle) {
        callback(msoStyle);
        return;
      }
    }

    callback(fallback);
  });
}

function resolveSenderEmail(item, callback) {
  var userProfile = Office.context.mailbox && Office.context.mailbox.userProfile;
  var fallbackEmail = (userProfile && userProfile.emailAddress ? userProfile.emailAddress : '').trim().toLowerCase();

  if (item && item.from && typeof item.from.getAsync === 'function') {
    item.from.getAsync(function (result) {
      var fromEmail =
        result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.emailAddress
          ? result.value.emailAddress.trim().toLowerCase()
          : '';
      callback(fromEmail || fallbackEmail);
    });
  } else {
    callback(fallbackEmail);
  }
}

function removeExistingSignatureAttachments(item, callback) {
  if (!item || typeof item.getAttachmentsAsync !== 'function' || typeof item.removeAttachmentAsync !== 'function') {
    callback();
    return;
  }

  item.getAttachmentsAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value || !result.value.length) {
      callback();
      return;
    }

    var toRemove = [];
    for (var i = 0; i < result.value.length; i++) {
      var att = result.value[i];
      if (att.isInline && Object.prototype.hasOwnProperty.call(SIG_FILE_NAMES, att.name)) {
        toRemove.push(att.id);
      }
    }

    if (toRemove.length === 0) {
      callback();
      return;
    }

    var remaining = toRemove.length;
    for (var j = 0; j < toRemove.length; j++) {
      item.removeAttachmentAsync(toRemove[j], function () {
        remaining--;
        if (remaining === 0) {
          callback();
        }
      });
    }
  });
}

function applySignature(event, isFromChanged) {
  var item = Office.context.mailbox && Office.context.mailbox.item;

  resolveSenderEmail(item, function (userEmail) {
    var person = Object.prototype.hasOwnProperty.call(STAFF, userEmail) ? STAFF[userEmail] : null;

    console.log(userEmail);

    // Unknown employee: on new compose exit cleanly; on From switch remove previous signature & attachments
    if (!person) {
      console.warn('No signature configured for:', userEmail || '(unknown user)');
      if (isFromChanged && item && item.body && typeof item.body.setSignatureAsync === 'function') {
        removeExistingSignatureAttachments(item, function () {
          item.body.setSignatureAsync('', { coercionType: Office.CoercionType.Html }, function () {
            event.completed();
          });
        });
      } else {
        event.completed();
      }
      return;
    }

    // Fixed 672x216 card: all dimensions divisible by 12 (exact Word inches at 96 DPI + crisp 125%/150%/200% scaling),
    // and all row heights >= 24px (above <address>'s ~17px fallback baseline on Classic Outlook Reply/Forward).
    // Widths: 12 (left pad) + 336 (info) + 60 (mid pad) + 264 (logo) = 672 (56 exact 12px stripes).
    // Heights: 84 (top pad + header) + 24 (phone) + 24 (web) + 36 (address + bottom pad) = 168, plus 48 (stripes) = 216.
    var cardWidth = 672;
    var base = 'https://mcmcoproperty.github.io/mcmcopropertyassets/v2/';
    var personBase = base + (person.folder || '');

    // Keep `name` as a plain filename (no slashes) and give every <img> its own unique cid: file
    // at exact aspect ratio so Gmail mobile replies never drop duplicate CIDs or rescale height:auto.
    var files = [
      { name: 'sig-stripes.png',  url: base + 'sig-stripes.png' },
      { name: 'sig-bg-left.png',  url: base + 'sig-bg-left.png' },
      { name: 'sig-bg-mid.png',   url: base + 'sig-bg-mid.png' },
      { name: 'sig-header.png',   url: personBase + 'sig-header.png' },
      { name: 'sig-phone.png',    url: personBase + 'sig-phone.png' },
      { name: 'sig-web.png',      url: base + 'sig-web.png' },
      { name: 'sig-address.png',  url: base + 'sig-address.png' },
      { name: 'sig-logo.png',     url: base + 'sig-logo.png' }
    ];

    // Default to external URLs; switch each to cid: once attached inline
    var srcMap = {};
    for (var i = 0; i < files.length; i++) {
      srcMap[files[i].name] = files[i].url;
    }

    var sigTextStyle = 'font-family:Aptos, Calibri, Helvetica, Arial, sans-serif; font-size:12pt; color:#000000;';

    function insertSignature() {
      // Combine flat colspan="4" table with <address> height/max-height/overflow:hidden clipping
      // and HTML align="top" on every <img> so Classic Outlook Reply/Forward preserves box height and kills descender gaps.
      var deadLinkStyle = 'display:block; border:none; text-decoration:none; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none; line-height:1px; font-size:1px;';
      var linkAnchorStyle = 'display:block; border:none; text-decoration:none; line-height:1px; font-size:1px;';

      function linkTdStyle(h) {
        return 'width:336px; min-width:336px; max-width:336px; height:' + h + 'px; padding:0; line-height:1px; font-size:1px;';
      }

      function linkAddrStyle(h) {
        return 'display:block; font-style:normal; margin:0; padding:0; width:336px; max-width:336px; height:' + h + 'px; max-height:' + h + 'px; overflow:hidden; line-height:1px; font-size:1px; white-space:nowrap;';
      }

      function deadAddrStyle(w, h) {
        return 'display:block; font-style:normal; margin:0; padding:0; width:' + w + 'px; max-width:' + w + 'px; height:' + h + 'px; max-height:' + h + 'px; overflow:hidden; line-height:1px; font-size:1px; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;';
      }

      var stripeImg =
        '<address style="' + deadAddrStyle(cardWidth, 48) + '"><a name="sig-stripes" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-stripes.png'] + '" width="' + cardWidth + '" height="48" border="0" align="top" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:' + cardWidth + 'px; min-width:' + cardWidth + 'px; max-width:' + cardWidth + 'px; height:48px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>';

      // Inner table gives each left-column slice its own fixed-height TD (336 x 168 total)
      var leftStackHtml = [
        '<table cellpadding="0" cellspacing="0" border="0" width="336" style="width:336px; min-width:336px; max-width:336px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">',
        '  <tr><td width="336" height="84" align="left" valign="top" style="width:336px; min-width:336px; max-width:336px; height:84px; padding:0; line-height:1px; font-size:1px;"><address style="' + deadAddrStyle(336, 84) + '"><a name="sig-header" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-header.png'] + '" width="336" height="84" border="0" align="top" alt="' + person.headerAlt + '" draggable="false" style="display:block; vertical-align:top; border:none; width:336px; min-width:336px; max-width:336px; height:84px; font-size:14px; line-height:20px; color:#4f1113; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address></td></tr>',
        '  <tr><td width="336" height="24" align="left" valign="top" style="' + linkTdStyle(24) + '"><address style="' + linkAddrStyle(24) + '"><a href="' + person.phoneHref + '" style="' + linkAnchorStyle + '"><img src="' + srcMap['sig-phone.png'] + '" width="336" height="24" border="0" align="top" alt="' + person.phoneAlt + '" style="display:block; vertical-align:top; border:none; width:336px; min-width:336px; max-width:336px; height:24px; font-size:12px; line-height:18px; color:#4f1113;"></a></address></td></tr>',
        '  <tr><td width="336" height="24" align="left" valign="top" style="' + linkTdStyle(24) + '"><address style="' + linkAddrStyle(24) + '"><a href="https://mcmanuscoproperty.com.au" style="' + linkAnchorStyle + '"><img src="' + srcMap['sig-web.png'] + '" width="336" height="24" border="0" align="top" alt="mcmanuscoproperty.com.au" style="display:block; vertical-align:top; border:none; width:336px; min-width:336px; max-width:336px; height:24px; font-size:12px; line-height:18px; color:#4f1113;"></a></address></td></tr>',
        '  <tr><td width="336" height="36" align="left" valign="top" style="' + linkTdStyle(36) + '"><address style="' + linkAddrStyle(36) + '"><a href="https://maps.google.com/?q=48+Sailors+Bay+Road,+Northbridge+NSW+2063" style="' + linkAnchorStyle + '"><img src="' + srcMap['sig-address.png'] + '" width="336" height="36" border="0" align="top" alt="48 Sailors Bay Road, Northbridge NSW 2063" style="display:block; vertical-align:top; border:none; width:336px; min-width:336px; max-width:336px; height:36px; font-size:12px; line-height:18px; color:#4f1113;"></a></address></td></tr>',
        '</table>'
      ].join('');

      // Single flat 2-row table with colspan="4" on Row 2 so Classic Outlook (Word) never flattens a wrapper table on Forward/Reply
      var signatureHtml = [
        '<div style="' + sigTextStyle + '"><br></div>',
        '<div style="' + sigTextStyle + '">Kind regards,</div>',
        '<div style="' + sigTextStyle + '"><br></div>',
        '<table cellpadding="0" cellspacing="0" border="0" width="' + cardWidth + '" style="width:' + cardWidth + 'px; min-width:' + cardWidth + 'px; max-width:' + cardWidth + 'px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; color:#4f1113; font-family:\'Century Gothic\', Helvetica, Arial, sans-serif; -webkit-text-size-adjust:100%;">',
        '  <tr>',
        // SECTION 1A: Fixed Left Padding Spacer (12 x 168)
        '    <td width="12" height="168" align="left" valign="top" style="width:12px; min-width:12px; max-width:12px; height:168px; padding:0; line-height:1px; font-size:1px;">',
        '      <address style="' + deadAddrStyle(12, 168) + '"><a name="sig-bg-left" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-bg-left.png'] + '" width="12" height="168" border="0" align="top" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:12px; min-width:12px; max-width:12px; height:168px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
        '    </td>',
        // SECTION 1B: Fixed Left Info Stacked in an Inner Table (336 x 168 total)
        '    <td width="336" height="168" align="left" valign="top" style="width:336px; min-width:336px; max-width:336px; height:168px; padding:0; line-height:1px; font-size:1px;">',
               leftStackHtml,
        '    </td>',
        // SECTION 2: Fixed Middle Spacer (60 x 168)
        '    <td width="60" height="168" align="left" valign="top" style="width:60px; min-width:60px; max-width:60px; height:168px; padding:0; line-height:1px; font-size:1px;">',
        '      <address style="' + deadAddrStyle(60, 168) + '"><a name="sig-bg-mid" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-bg-mid.png'] + '" width="60" height="168" border="0" align="top" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:60px; min-width:60px; max-width:60px; height:168px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
        '    </td>',
        // SECTION 3: Fixed Right Logo + Margin (264 x 168)
        '    <td width="264" height="168" align="left" valign="top" style="width:264px; min-width:264px; max-width:264px; height:168px; padding:0; line-height:1px; font-size:1px;">',
        '      <address style="display:block; font-style:normal; margin:0; padding:0; width:264px; max-width:264px; height:168px; max-height:168px; overflow:hidden; line-height:1px; font-size:1px; white-space:nowrap;"><a href="https://mcmanuscoproperty.com.au" style="' + linkAnchorStyle + '"><img src="' + srcMap['sig-logo.png'] + '" width="264" height="168" border="0" align="top" alt="McManus &amp; Co Property" style="display:block; vertical-align:top; border:none; width:264px; min-width:264px; max-width:264px; height:168px; font-size:14px; line-height:20px; color:#4f1113;"></a></address>',
        '    </td>',
        '  </tr>',
        '  <tr>',
        '    <td colspan="4" width="' + cardWidth + '" height="48" align="left" valign="top" style="width:' + cardWidth + 'px; min-width:' + cardWidth + 'px; max-width:' + cardWidth + 'px; height:48px; padding:0; line-height:1px; font-size:1px;">',
               stripeImg,
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

    function attachAllAndInsert() {
      var canUrlAttach = item && typeof item.addFileAttachmentAsync === 'function';
      var canBase64Attach = item && typeof item.addFileAttachmentFromBase64Async === 'function';

      if (!canUrlAttach && !canBase64Attach) {
        insertSignature();
        return;
      }

      var platform = Office.context.diagnostics && Office.context.diagnostics.platform;
      var isMobile =
        typeof Office.PlatformType !== 'undefined' &&
        (platform === Office.PlatformType.iOS || platform === Office.PlatformType.Android);

      var remaining = files.length;

      function onFileDone() {
        remaining--;
        if (remaining === 0) {
          insertSignature();
        }
      }

      function attachViaBase64(file) {
        if (!canBase64Attach || typeof XMLHttpRequest === 'undefined') {
          onFileDone();
          return;
        }

        try {
          var xhr = new XMLHttpRequest();
          xhr.open('GET', file.url, true);
          xhr.responseType = 'arraybuffer';
          xhr.onload = function () {
            if (xhr.status >= 200 && xhr.status < 300 && xhr.response) {
              var base64 = arrayBufferToBase64(xhr.response);
              item.addFileAttachmentFromBase64Async(
                base64,
                file.name,
                { isInline: true },
                function (result) {
                  if (result.status === Office.AsyncResultStatus.Succeeded) {
                    srcMap[file.name] = 'cid:' + file.name;
                  } else {
                    console.error('addFileAttachmentFromBase64Async failed for ' + file.name + ':', result.error.message);
                  }
                  onFileDone();
                }
              );
            } else {
              console.error('Image fetch failed for ' + file.name + ': HTTP ' + xhr.status);
              onFileDone();
            }
          };
          xhr.onerror = function () {
            console.error('Network error fetching ' + file.name);
            onFileDone();
          };
          xhr.send();
        } catch (err) {
          console.error('Base64 attach error for ' + file.name + ':', err);
          onFileDone();
        }
      }

      function attachFile(file) {
        if (isMobile || !canUrlAttach) {
          attachViaBase64(file);
          return;
        }

        item.addFileAttachmentAsync(
          file.url,
          file.name,
          { isInline: true },
          function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              srcMap[file.name] = 'cid:' + file.name;
              onFileDone();
            } else if (canBase64Attach) {
              attachViaBase64(file);
            } else {
              console.error('addFileAttachmentAsync failed for ' + file.name + ':', result.error.message);
              onFileDone();
            }
          }
        );
      }

      for (var j = 0; j < files.length; j++) {
        attachFile(files[j]);
      }
    }

    detectDefaultTextStyle(item, function (detectedStyle) {
      sigTextStyle = detectedStyle;
      if (isFromChanged) {
        removeExistingSignatureAttachments(item, attachAllAndInsert);
      } else {
        attachAllAndInsert();
      }
    });
  });
}

function onNewMessageComposeHandler(event) {
  applySignature(event, false);
}

function onMessageFromChangedHandler(event) {
  applySignature(event, true);
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
Office.actions.associate('onMessageFromChangedHandler', onMessageFromChangedHandler);
