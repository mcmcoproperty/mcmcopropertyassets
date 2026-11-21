Office.onReady();

// Per-employee directory keyed by lowercase Microsoft 365 email address.
// Shared assets (stripes, background, web, address, logo) stay in the root folder.
// For future staff, set folder to e.g. 'staff/james/' and upload their sig-header.png and sig-phone.png there.
var STAFF = {
  'louise@mcmanuscoproperty.com': {
    folder: 'staff/louise/',
    headerAlt: 'Louise Craighill - Chief Financial Officer',
    phoneHref: 'tel:+61422726186',
    phoneAlt: '+61 422 726 186'
  },
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
  }
};

var SIG_FILE_NAMES = {
  'sig-stripes.png': true,
  'sig-background.png': true,
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

    // sig-stripes.png is 3000x1 (native 12px stripes).
    // Keep imgWidth = 3000 for exact 12px stripes (or scale proportionally, e.g. 2500 for 10px stripes).
    var imgWidth = 3000;
    var base = 'https://mcmcoproperty.github.io/mcmcopropertyassets/v1/';
    var personBase = base + (person.folder || '');

    // Keep `name` as a plain filename (no slashes) for valid cid: references
    var files = [
      { name: 'sig-stripes.png',     url: base + 'sig-stripes.png' },
      { name: 'sig-background.png',  url: base + 'sig-background.png' },
      { name: 'sig-header.png',      url: personBase + 'sig-header.png' },
      { name: 'sig-phone.png',       url: personBase + 'sig-phone.png' },
      { name: 'sig-web.png',         url: base + 'sig-web.png' },
      { name: 'sig-address.png',     url: base + 'sig-address.png' },
      { name: 'sig-logo.png',        url: base + 'sig-logo.png' }
    ];

    // Default to external URLs; switch each to cid: once attached inline
    var srcMap = {};
    for (var i = 0; i < files.length; i++) {
      srcMap[files[i].name] = files[i].url;
    }

    function insertSignature() {
      var deadLinkStyle = 'display:block; border:none; text-decoration:none; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none; line-height:0; font-size:0;';
      var linkRowStyle = 'display:block; font-style:normal; margin:0; padding:0; width:335px; max-width:335px; height:26px; max-height:26px; overflow:hidden; line-height:0; font-size:0; white-space:nowrap;';

      var stripeImg =
        '<a name="sig-stripes" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-stripes.png'] + '" width="' + imgWidth + '" height="40" alt="" draggable="false" style="display:block; border:none; width:' + imgWidth + 'px; min-width:' + imgWidth + 'px; max-width:none; height:40px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a>';

      var leftStackHtml = [
        '<address style="display:block; font-style:normal; margin:0; padding:0; width:335px; max-width:335px; height:15px; max-height:15px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-top" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-background.png'] + '" width="335" height="15" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:15px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
        '<address style="display:block; font-style:normal; margin:0; padding:0; width:335px; max-width:335px; height:69px; max-height:69px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-header" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-header.png'] + '" width="335" height="69" alt="' + person.headerAlt + '" draggable="false" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:69px; font-size:16px; line-height:22px; color:#efdfd4; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
        '<address style="' + linkRowStyle + '"><a href="' + person.phoneHref + '" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;"><img src="' + srcMap['sig-phone.png'] + '" width="335" height="26" alt="' + person.phoneAlt + '" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:26px; font-size:12px; line-height:26px; color:#efdfd4;"></a></address>',
        '<address style="' + linkRowStyle + '"><a href="https://mcmanuscoproperty.com.au" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;"><img src="' + srcMap['sig-web.png'] + '" width="335" height="26" alt="mcmanuscoproperty.com.au" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:26px; font-size:12px; line-height:26px; color:#efdfd4;"></a></address>',
        '<address style="' + linkRowStyle + '"><a href="https://maps.google.com/?q=48+Sailors+Bay+Road,+Northbridge+NSW+2063" style="display:block; border:none; text-decoration:none; line-height:0; font-size:0;"><img src="' + srcMap['sig-address.png'] + '" width="335" height="26" alt="48 Sailors Bay Road, Northbridge NSW 2063" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:26px; font-size:12px; line-height:26px; color:#efdfd4;"></a></address>',
        '<address style="display:block; font-style:normal; margin:0; padding:0; width:335px; max-width:335px; height:8px; max-height:8px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-bottom" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-background.png'] + '" width="335" height="8" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:335px; min-width:335px; max-width:335px; height:8px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>'
      ].join('');

      // Split 170px into 3 strips (57 + 57 + 56) so each stays under Gmail's image-attachment height threshold
      var midStackHtml = [
        '<address style="display:block; font-style:normal; margin:0; padding:0; width:100%; height:57px; max-height:57px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-mid-1" draggable="false" style="width:100%; ' + deadLinkStyle + '"><img src="' + srcMap['sig-background.png'] + '" width="60" height="57" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:100%; min-width:60px; height:57px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
        '<address style="display:block; font-style:normal; margin:0; padding:0; width:100%; height:57px; max-height:57px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-mid-2" draggable="false" style="width:100%; ' + deadLinkStyle + '"><img src="' + srcMap['sig-background.png'] + '" width="60" height="57" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:100%; min-width:60px; height:57px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>',
        '<address style="display:block; font-style:normal; margin:0; padding:0; width:100%; height:56px; max-height:56px; overflow:hidden; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"><a name="sig-bg-mid-3" draggable="false" style="width:100%; ' + deadLinkStyle + '"><img src="' + srcMap['sig-background.png'] + '" width="60" height="56" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:100%; min-width:60px; height:56px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a></address>'
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
        '              <a name="sig-bg-left" draggable="false" style="' + deadLinkStyle + '"><img src="' + srcMap['sig-background.png'] + '" width="15" height="170" alt="" draggable="false" style="display:block; vertical-align:top; border:none; width:15px; min-width:15px; max-width:15px; height:170px; cursor:default; pointer-events:none; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;"></a>',
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
        '            <address style="display:block; font-style:normal; margin:0; padding:0; width:260px; max-width:260px; height:170px; max-height:170px; overflow:hidden; background-color:#4f1113; line-height:0; font-size:0; white-space:nowrap;">',
        '              <a href="https://mcmanuscoproperty.com.au" style="display:block; border:none; text-decoration:none; background-color:#4f1113; line-height:0; font-size:0;">',
        '                <img src="' + srcMap['sig-logo.png'] + '" width="260" height="170" alt="McManus &amp; Co Property" style="display:block; vertical-align:top; border:none; width:260px; min-width:260px; max-width:260px; height:170px; background-color:#4f1113; font-size:14px; line-height:20px; color:#efdfd4;">',
        '              </a>',
        '            </address>',
        '          </td>',
        '        </tr>',
        '      </table>',
        '    </td>',
        '  </tr>',
        '  <tr>',
        '    <td width="100%" height="40" bgcolor="#4f1113" style="width:100%; height:40px; max-height:40px; background-color:#4f1113; padding:0; line-height:0; font-size:0;">',
        '      <address style="display:block; font-style:normal; margin:0; padding:0; width:670px; min-width:100%; height:40px; max-height:40px; overflow:hidden; background-color:#4f1113; line-height:0; font-size:0; pointer-events:none; cursor:default; -webkit-user-drag:none; user-select:none; -webkit-user-select:none;">',
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

    if (isFromChanged) {
      removeExistingSignatureAttachments(item, attachAllAndInsert);
    } else {
      attachAllAndInsert();
    }
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
