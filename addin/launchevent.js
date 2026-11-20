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
      '
