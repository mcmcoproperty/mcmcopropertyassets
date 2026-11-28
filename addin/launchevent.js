Office.onReady();

// Per-employee directory keyed by lowercase Microsoft 365 email address.
// Shared assets (stripes, bg-pad, web, address, logo) stay in the root folder.
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

    // Fixed 672x216 card: all dimensions divisible by 4 for crisp 125%/150%/200% DPI scaling.
    // Widths: 16 (left pad) + 336 (info) + 60 (mid pad) + 260 (logo) = 672 (56 exact 12px stripes).
    // Heights: 16 (top) + 68 (header) + 28 + 28 + 28 (links) + 8 (bottom) = 176, plus 40 (stripes) = 216.
    var cardWidth = 672;
    var base = 'https://mcmcoproperty.github.io/mcmcopropertyassets/v1/';
    var personBase = base + (person.folder || '');

    // Keep `name` as a plain filename (no slashes) for valid cid: references
    var files = [
      { name: 'sig-stripes.png',  url: base + 'sig-stripes.png' },
      { name: 'sig-bg-pad.png',   url: base + 'sig-bg-pad.png' },
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
      // Use line-height:1px; font-size:1px; without mso-line-height-rule:exactly so Outlook Mobile stripping
      // mso-* rules never introduces descender gaps in Classic Outlook, while preventing line-height:0 clipping.
      var deadLinkStyle = 'display:block; border:none; text-decoration:none; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none; line-height:1px; font-size:1px;';
      var linkRowTdStyle = 'width:336px; min-width:336px; max-width:336px; height:28px; padding:0; line-height:1px; font-size:1px;';
      var linkRowStyle = 'display:block; font-style:normal; margin:0; padding:0; width:336px; max-width:336px; height:28px; max-height:28px; overflow:hidden; line-height:1px; font-size:1px; white-space:nowrap;';

      var stripeImg =
        '<a name="sig-stripes" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-stripes.png'] + '" width="' + cardWidth + '" height="40" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:' + cardWidth + 'px; min-width:' + cardWidth + 'px; max-width:' + cardWidth + 'px; height:40px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a>';

      // Inner table gives each left-column slice its own fixed-height TD (336 x 176 total)
      var leftStackHtml = [
        '<table cellpadding="0" cellspacing="0" border="0" width="336" style="width:336px; min-width:336px; max-width:336px; table-layout:fixed; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">',
        '  <tr><td width="336" height="16" valign="top" style="width:336px; min-width:336px; max-width:336px; height:16px; padding:0; line-height:1px; font-size:1px;"><address style="display:block; font-style:normal; margin:0; padding:0; width:336px; max-width:336px; height:16px; max-height:16px; overflow:hidden; line-height:1px; font-size:1px; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-top" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-bg-pad.png'] + '" width="336" height="16" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:336px; min-width:336px; max-width:336px; height:16px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address></td></tr>',
        '  <tr><td width="336" height="68" valign="top" style="width:336px; min-width:336px; max-width:336px; height:68px; padding:0; line-height:1px; font-size:1px;"><address style="display:block; font-style:normal; margin:0; padding:0; width:336px; max-width:336px; height:68px; max-height:68px; overflow:hidden; line-height:1px; font-size:1px; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-header" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-header.png'] + '" width="336" height="68" alt="' + person.headerAlt + '" draggable="false" style="display:block; vertical-align:top; border:none; width:336px; min-width:336px; max-width:336px; height:68px; font-size:16px; line-height:22px; color:#4f1113; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address></td></tr>',
        '  <tr><td width="336" height="28" valign="top" style="' + linkRowTdStyle + '"><address style="' + linkRowStyle + '"><a href="' + person.phoneHref + '" style="display:block; border:none; text-decoration:none; line-height:1px; font-size:1px;"><img src="' + srcMap['sig-phone.png'] + '" width="336" height="28" alt="' + person.phoneAlt + '" style="display:block; vertical-align:top; border:none; width:336px; min-width:336px; max-width:336px; height:28px; font-size:12px; line-height:28px; color:#4f1113;"></a></address></td></tr>',
        '  <tr><td width="336" height="28" valign="top" style="' + linkRowTdStyle + '"><address style="' + linkRowStyle + '"><a href="https://mcmanuscoproperty.com.au" style="display:block; border:none; text-decoration:none; line-height:1px; font-size:1px;"><img src="' + srcMap['sig-web.png'] + '" width="336" height="28" alt="mcmanuscoproperty.com.au" style="display:block; vertical-align:top; border:none; width:336px; min-width:336px; max-width:336px; height:28px; font-size:12px; line-height:28px; color:#4f1113;"></a></address></td></tr>',
        '  <tr><td width="336" height="28" valign="top" style="' + linkRowTdStyle + '"><address style="' + linkRowStyle + '"><a href="https://maps.google.com/?q=48+Sailors+Bay+Road,+Northbridge+NSW+2063" style="display:block; border:none; text-decoration:none; line-height:1px; font-size:1px;"><img src="' + srcMap['sig-address.png'] + '" width="336" height="28" alt="48 Sailors Bay Road, Northbridge NSW 2063" style="display:block; vertical-align:top; border:none; width:336px; min-width:336px; max-width:336px; height:28px; font-size:12px; line-height:28px; color:#4f1113;"></a></address></td></tr>',
        '  <tr><td width="336" height="8" valign="top" style="width:336px; min-width:336px; max-width:336px; height:8px; padding:0; line-height:1px; font-size:1px;"><address style="display:block; font-style:normal; margin:0; padding:0; width:336px; max-width:336px; height:8px; max-height:8px; overflow:hidden; line-height:1px; font-size:1px; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-bottom" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-bg-pad.png'] + '" width="336" height="8" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:336px; min-width:336px; max-width:336px; height:8px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address></td></tr>',
        '</table>'
      ].join('');

      var signatureHtml = [
        '<div style="' + sigTextStyle + '"><br></div>',
        '<div style="' + sigTextStyle + '">Kind regards,</div>',
        '<div style="' + sigTextStyle + '"><br></div>',
        '<table cellpadding="0" cellspacing="0" border="0" width="' + cardWidth + '" style="width:' + cardWidth + 'px; min-width:' + cardWidth + 'px; max-width:' + cardWidth + 'px; table-layout:fixed; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; color:#4f1113; font-family:\'Century Gothic\', Helvetica, Arial, sans-serif; -webkit-text-size-adjust:100%;">',
        '  <tr>',
        '    <td width="' + cardWidth + '" height="176" align="left" valign="top" style="width:' + cardWidth + 'px; min-width:' + cardWidth + 'px; max-width:' + cardWidth + 'px; height:176px; padding:0; font-size:0;">',
        '      <table cellpadding="0" cellspacing="0" border="0" width="' + cardWidth + '" style="width:' + cardWidth + 'px; min-width:' + cardWidth + 'px; max-width:' + cardWidth + 'px; table-layout:fixed; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; font-size:0;">',
        '        <tr>',
        // SECTION 1A: Fixed Left Padding Spacer (16 x 176)
        '          <td width="16" height="176" valign="top" style="width:16px; min-width:16px; max-width:16px; height:176px; padding:0; line-height:1px; font-size:1px;">',
        '            <address style="display:block; font-style:normal; margin:0; padding:0; width:16px; max-width:16px; height:176px; max-height:176px; overflow:hidden; line-height:1px; font-size:1px; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-left" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-bg-pad.png'] + '" width="16" height="176" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:16px; min-width:16px; max-width:16px; height:176px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
        '          </td>',
        // SECTION 1B: Fixed Left Info Stacked in an Inner Table (336 x 176 total)
        '          <td width="336" height="176" valign="top" style="width:336px; min-width:336px; max-width:336px; height:176px; padding:0; font-size:0;">',
                     leftStackHtml,
        '          </td>',
        // SECTION 2: Fixed Middle Spacer (60 x 176)
        '          <td width="60" height="176" valign="top" style="width:60px; min-width:60px; max-width:60px; height:176px; padding:0; line-height:1px; font-size:1px;">',
        '            <address style="display:block; font-style:normal; margin:0; padding:0; width:60px; max-width:60px; height:176px; max-height:176px; overflow:hidden; line-height:1px; font-size:1px; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-mid" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-bg-pad.png'] + '" width="60" height="176" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:60px; min-width:60px; max-width:60px; height:176px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
        '          </td>',
        // SECTION 3: Fixed Right Logo + Margin (260 x 176)
        '          <td width="260" height="176" valign="top" style="width:260px; min-width:260px; max-width:260px; height:176px; padding:0; line-height:1px; font-size:1px;">',
        '            <address style="display:block; font-style:normal; margin:0; padding:0; width:260px; max-width:260px; height:176px; max-height:176px; overflow:hidden; line-height:1px; font-size:1px; white-space:nowrap;"><a href="https://mcmanuscoproperty.com.au" style="display:block; border:none; text-decoration:none; line-height:1px; font-size:1px;"><img src="' + srcMap['sig-logo.png'] + '" width="260" height="176" alt="McManus &amp; Co Property" style="display:block; vertical-align:top; border:none; width:260px; min-width:260px; max-width:260px; height:176px; font-size:14px; line-height:20px; color:#4f1113;"></a></address>',
        '          </td>',
        '        </tr>',
        '      </table>',
        '    </td>',
        '  </tr>',
        '  <tr>',
        '    <td width="' + cardWidth + '" height="40" align="left" valign="top" style="width:' + cardWidth + 'px; min-width:' + cardWidth + 'px; max-width:' + cardWidth + 'px; height:40px; max-height:40px; padding:0; line-height:1px; font-size:1px;">',
        '      <address style="display:block; font-style:normal; margin:0; padding:0; width:' + cardWidth + 'px; max-width:' + cardWidth + 'px; height:40px; max-height:40px; overflow:hidden; line-height:1px; font-size:1px; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;">' + stripeImg + '</address>',
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
