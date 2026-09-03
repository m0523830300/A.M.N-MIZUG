/* AMN live action status - isolated from page logic */
(function () {
  'use strict';
  if (window.__amnStatusReady) return;
  window.__amnStatusReady = true;

  var style = document.createElement('style');
  style.textContent =
    '#amn-action-status{position:fixed;z-index:999999;left:50%;bottom:max(18px,env(safe-area-inset-bottom));' +
    'transform:translate(-50%,14px);opacity:0;pointer-events:none;min-width:180px;max-width:calc(100vw - 28px);' +
    'padding:11px 18px;border-radius:999px;background:#153E6F;color:#fff;text-align:center;' +
    'font-family:Heebo,Arial,sans-serif;font-size:14px;font-weight:700;box-shadow:0 8px 28px rgba(11,42,80,.25);' +
    'transition:opacity .18s ease,transform .18s ease}' +
    '#amn-action-status.amn-show{opacity:1;transform:translate(-50%,0)}' +
    '#amn-action-status.amn-success{background:#178A4C}' +
    '#amn-action-status.amn-error{background:#B7352A}' +
    '#amn-action-status .amn-spin{display:inline-block;width:13px;height:13px;margin-left:8px;vertical-align:-2px;' +
    'border:2px solid rgba(255,255,255,.45);border-top-color:#fff;border-radius:50%;animation:amn-spin .7s linear infinite}' +
    '@keyframes amn-spin{to{transform:rotate(360deg)}}';
  document.head.appendChild(style);

  var box = document.createElement('div');
  box.id = 'amn-action-status';
  box.setAttribute('role', 'status');
  box.setAttribute('aria-live', 'polite');
  document.body.appendChild(box);

  var timer = 0;
  function show(message, state, delay) {
    clearTimeout(timer);
    box.className = 'amn-show amn-' + state;
    box.innerHTML = (state === 'loading' ? '<span class="amn-spin"></span>' : '') +
      String(message || '');
    if (delay) timer = setTimeout(function () { box.className = ''; }, delay);
  }

  window.showAppStatus = function (message, state) {
    show(message, state || 'success', state === 'error' ? 4000 : 1800);
  };

  var labels = {
    uploadFile: ['מעלה קובץ…', 'הקובץ הועלה ✓'],
    uploadOwnerPhoto: ['מעלה קובץ…', 'הקובץ הועלה ✓'],
    recordPayment: ['שומר תשלום…', 'התשלום נשמר ✓'],
    notifyCustomerOnWay: ['שולח הודעה…', 'ההודעה נשלחה ✓'],
    updateStatus: ['מעדכן סטטוס…', 'הסטטוס עודכן ✓'],
    createCall: ['פותח קריאה…', 'הקריאה נפתחה ✓'],
    saveCall: ['שומר קריאה…', 'הקריאה נשמרה ✓'],
    checkIn: ['שומר כניסה…', 'הכניסה נשמרה ✓'],
    checkOut: ['שומר יציאה…', 'היציאה נשמרה ✓'],
    saveAttendance: ['שומר דיווח…', 'הדיווח נשמר ✓'],
    saveWork: ['שומר עבודה…', 'העבודה נשמרה ✓'],
    saveTask: ['שומר משימה…', 'המשימה נשמרה ✓'],
    completeTask: ['מסמן כהושלם…', 'המשימה הושלמה ✓']
  };

  function requestInfo(input, init) {
    var method = String((init && init.method) || 'GET').toUpperCase();
    var action = '';
    if (init && typeof init.body === 'string' && init.body.charAt(0) === '{') {
      try { action = JSON.parse(init.body).action || ''; } catch (ignore) {}
    }
    var isRead = method === 'GET' || /^(get|list|load|login|validate|check)/i.test(action);
    return {
      labels: labels[action] || (isRead ?
        ['טוען נתונים…', 'הנתונים עודכנו ✓'] :
        ['מבצע פעולה…', 'הפעולה הושלמה ✓'])
    };
  }

  var originalFetch = window.fetch.bind(window);
  window.fetch = function (input, init) {
    var info = requestInfo(input, init || {});
    show(info.labels[0], 'loading', 0);

    return originalFetch(input, init).then(function (response) {
      response.clone().json().then(function (payload) {
        if (payload && payload.success === false) {
          show(payload.error || 'הפעולה לא הושלמה', 'error', 4000);
        } else {
          show(info.labels[1], 'success', 1800);
        }
      }).catch(function () {
        if (response.ok) show(info.labels[1], 'success', 1800);
      });
      if (!response.ok) show('אירעה תקלה. אפשר לנסות שוב.', 'error', 4000);
      return response;
    }).catch(function (error) {
      show('אין תקשורת עם השרת. אפשר לנסות שוב.', 'error', 4000);
      throw error;
    });
  };

  window.addEventListener('offline', function () {
    show('אין חיבור לאינטרנט', 'error', 4000);
  });
  window.addEventListener('online', function () {
    show('החיבור חזר ✓', 'success', 1800);
  });
})();
