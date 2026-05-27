/**
 * globe-mount.js
 * Inietta <globe-loader> in ogni div.globe-host appena React lo aggiunge al DOM.
 * Legge dimensione e colori da CSS custom properties impostate inline.
 */
(function () {
  function mountGlobes() {
    document.querySelectorAll('.globe-host:not([data-gl-mounted])').forEach(function (el) {
      el.setAttribute('data-gl-mounted', '1');

      var cs   = getComputedStyle(el);
      var size = (cs.getPropertyValue('--gl-size') || '').trim() || '120';
      var ink  = (cs.getPropertyValue('--gl-ink')  || '').trim() || '#1a1a1a';
      var bg   = (cs.getPropertyValue('--gl-bg')   || '').trim() || '#f4f1ec';

      var globe = document.createElement('globe-loader');
      globe.setAttribute('size', size);
      globe.style.setProperty('--gl-ink', ink);
      globe.style.setProperty('--gl-bg',  bg);

      el.appendChild(globe);
    });
  }

  /* Osserva tutte le modifiche al DOM (Dash/React ri-renderizza in modo dinamico) */
  var observer = new MutationObserver(mountGlobes);
  observer.observe(document.documentElement, { childList: true, subtree: true });

  /* Fallback per il primo caricamento */
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', mountGlobes);
  } else {
    mountGlobes();
  }
})();
