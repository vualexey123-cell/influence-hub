/* Influence Hub — cookie consent + consent-gated Google Analytics */
(function () {
  var GA_ID = 'G-KELKMRNNYP';
  var KEY = 'ih-consent';
  var PRIVACY_URL = '/privacy.html';

  function loadGA() {
    if (window.__ihGaLoaded) return;
    window.__ihGaLoaded = true;
    var s = document.createElement('script');
    s.async = true;
    s.src = 'https://www.googletagmanager.com/gtag/js?id=' + GA_ID;
    document.head.appendChild(s);
    window.dataLayer = window.dataLayer || [];
    window.gtag = function () { dataLayer.push(arguments); };
    gtag('js', new Date());
    gtag('config', GA_ID, { anonymize_ip: true });
  }

  function getChoice() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
  function setChoice(v) { try { localStorage.setItem(KEY, v); } catch (e) {} }

  var choice = getChoice();
  if (choice === 'accepted') { loadGA(); return; }
  if (choice === 'declined') { return; }

  function ready(fn) {
    if (document.readyState !== 'loading') fn();
    else document.addEventListener('DOMContentLoaded', fn);
  }

  ready(function () {
    var style = document.createElement('style');
    style.textContent =
      '#ih-cookie{position:fixed;left:16px;right:16px;bottom:16px;z-index:200;max-width:560px;margin:0 auto;' +
      'background:#1B1A17;color:#F5EFE4;border-radius:16px;padding:18px 20px;' +
      'box-shadow:0 18px 44px -18px rgba(0,0,0,.6);font-family:Manrope,system-ui,sans-serif;font-size:14px;line-height:1.5}' +
      '#ih-cookie a{color:#E8724F;text-decoration:none}#ih-cookie a:hover{text-decoration:underline}' +
      '#ih-cookie .ih-row{display:flex;gap:10px;margin-top:14px;flex-wrap:wrap}' +
      '#ih-cookie button{cursor:pointer;border:0;border-radius:999px;font:inherit;font-weight:600;padding:10px 18px}' +
      '#ih-cookie .ih-acc{background:#C13F22;color:#FFF9F2}' +
      '#ih-cookie .ih-dec{background:transparent;color:#B5AC9F;border:1px solid #3a372f}';
    document.head.appendChild(style);

    var b = document.createElement('div');
    b.id = 'ih-cookie';
    b.setAttribute('role', 'dialog');
    b.setAttribute('aria-label', 'Cookie consent');
    b.innerHTML =
      '<div>We use cookies to measure traffic and improve this website. ' +
      'Analytics cookies are set only if you accept. See our <a href="' + PRIVACY_URL + '">Privacy Policy</a>.</div>' +
      '<div class="ih-row"><button class="ih-acc" id="ih-acc">Accept</button>' +
      '<button class="ih-dec" id="ih-dec">Decline</button></div>';
    document.body.appendChild(b);

    document.getElementById('ih-acc').onclick = function () { setChoice('accepted'); b.remove(); loadGA(); };
    document.getElementById('ih-dec').onclick = function () { setChoice('declined'); b.remove(); };
  });
})();
