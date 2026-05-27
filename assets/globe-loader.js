// Globe Loader — componente riutilizzabile (Web Component).
// Uso:
//   <globe-loader></globe-loader>                 -> 200x200 inline
//   <globe-loader size="120"></globe-loader>      -> dimensione custom
//   GlobeLoader.showOverlay()  / .hideOverlay()   -> overlay fullscreen

(function () {
  const TEMPLATE = (size) => `
    <style>
      :host { display: inline-block; }
      .stage {
        width: ${size}px; height: ${size}px;
        position: relative; display: flex;
        align-items: center; justify-content: center;
      }
      .whirl, .whirl-inner { position: absolute; inset: 0; pointer-events: none; }
      .whirl-inner { inset: ${Math.round(size * 0.03)}px; }
      .whirl svg, .whirl-inner svg { width: 100%; height: 100%; }
      .whirl svg { animation: gl-cw 2.4s linear infinite; }
      .whirl-inner svg { animation: gl-ccw 1.8s linear infinite; }
      @keyframes gl-cw  { to { transform: rotate(360deg); } }
      @keyframes gl-ccw { to { transform: rotate(-360deg); } }
      .globe-wrap {
        width: ${Math.round(size * 0.7)}px;
        height: ${Math.round(size * 0.7)}px;
      }
      .ocean { fill: var(--gl-bg, #f4f1ec); }
      .land  { fill: var(--gl-ink, #1a1a1a); stroke: var(--gl-bg, #f4f1ec); stroke-width: 0.4; }
      .graticule { stroke: var(--gl-bg, #f4f1ec); stroke-width: 0.4; fill: none; opacity: 0.55; }
      .globe-outline { fill: none; stroke: var(--gl-ink, #1a1a1a); stroke-width: 1.2; }
    </style>
    <div class="stage">
      <div class="whirl">
        <svg viewBox="0 0 200 200">
          <g fill="none" stroke="var(--gl-ink, #1a1a1a)" stroke-linecap="round">
            <path d="M 100 6  A 94 94 0 0 1 188 60" stroke-width="1.6"/>
            <path d="M 14 130 A 94 94 0 0 1 40 178" stroke-width="1.2" opacity="0.55"/>
            <path d="M 100 194 A 94 94 0 0 1 30 160" stroke-width="0.9" opacity="0.35"/>
            <circle cx="194" cy="100" r="2.2" fill="var(--gl-ink, #1a1a1a)"/>
            <circle cx="172" cy="46"  r="1.4" fill="var(--gl-ink, #1a1a1a)" opacity="0.7"/>
            <circle cx="6"   cy="100" r="1.6" fill="var(--gl-ink, #1a1a1a)" opacity="0.5"/>
          </g>
        </svg>
      </div>
      <div class="whirl-inner">
        <svg viewBox="0 0 188 188">
          <g fill="none" stroke="var(--gl-ink, #1a1a1a)" stroke-linecap="round">
            <path d="M 94 6   A 88 88 0 0 1 168 50" stroke-width="0.9" opacity="0.6" stroke-dasharray="2 5"/>
            <path d="M 20 130 A 88 88 0 0 1 50 172" stroke-width="0.9" opacity="0.45" stroke-dasharray="2 5"/>
            <path d="M 94 182 A 88 88 0 0 1 22 140" stroke-width="0.7" opacity="0.3"  stroke-dasharray="1 4"/>
            <circle cx="94" cy="2" r="1.6" fill="var(--gl-ink, #1a1a1a)"/>
            <circle cx="180" cy="94" r="1.2" fill="var(--gl-ink, #1a1a1a)" opacity="0.6"/>
          </g>
        </svg>
      </div>
      <div class="globe-wrap">
        <svg class="globe" viewBox="-70 -70 140 140">
          <defs>
            <clipPath id="gl-clip-${size}"><circle cx="0" cy="0" r="66"/></clipPath>
          </defs>
          <circle cx="0" cy="0" r="66" class="ocean"/>
          <g clip-path="url(#gl-clip-${size})">
            <g class="land-layer"></g>
            <g class="grat-layer"></g>
          </g>
          <circle cx="0" cy="0" r="66" class="globe-outline"/>
        </svg>
      </div>
    </div>
  `;

  // Shared world topology (caricato una sola volta)
  let worldPromise = null;
  function getWorld() {
    if (worldPromise) return worldPromise;
    worldPromise = Promise.all([
      import('https://cdn.jsdelivr.net/npm/d3@7.9.0/+esm'),
      import('https://cdn.jsdelivr.net/npm/topojson-client@3.1.0/+esm'),
      fetch('https://cdn.jsdelivr.net/npm/world-atlas@2.0.2/countries-110m.json').then(r => r.json())
    ]).then(([d3, topo, atlas]) => ({
      d3, topo,
      features: topo.feature(atlas, atlas.objects.countries)
    }));
    return worldPromise;
  }

  class GlobeLoader extends HTMLElement {
    connectedCallback() {
      const size = parseInt(this.getAttribute('size') || '200', 10);
      const root = this.attachShadow({ mode: 'open' });
      root.innerHTML = TEMPLATE(size);
      this._size = size;
      this._raf = null;
      this._start = performance.now();
      getWorld().then((ctx) => this._init(ctx));
    }
    disconnectedCallback() {
      if (this._raf) cancelAnimationFrame(this._raf);
    }
    _init({ d3, features }) {
      const root = this.shadowRoot;
      const projection = d3.geoOrthographic().scale(66).translate([0, 0]).clipAngle(90);
      const path = d3.geoPath(projection);
      const graticule = d3.geoGraticule10();
      const landEl = root.querySelector('.land-layer');
      const gratEl = root.querySelector('.grat-layer');
      const tilt = -18, spin = 28;

      const tick = (now) => {
        const t = (now - this._start) / 1000;
        projection.rotate([(t * spin) % 360 - 180, tilt, 0]);
        let d = '';
        for (const f of features.features) {
          const p = path(f); if (p) d += p;
        }
        landEl.innerHTML = `<path d="${d}" class="land"/>`;
        const g = path(graticule);
        gratEl.innerHTML = g ? `<path d="${g}" class="graticule"/>` : '';
        this._raf = requestAnimationFrame(tick);
      };
      this._raf = requestAnimationFrame(tick);
    }
  }

  customElements.define('globe-loader', GlobeLoader);

  // Helper: overlay fullscreen
  let _overlay = null;
  GlobeLoader.showOverlay = function (opts = {}) {
    if (_overlay) return;
    _overlay = document.createElement('div');
    _overlay.style.cssText = `
      position: fixed; inset: 0; z-index: 9999;
      background: ${opts.background || 'rgba(244,241,236,0.92)'};
      backdrop-filter: blur(2px);
      display: flex; align-items: center; justify-content: center;
      opacity: 0; transition: opacity 200ms ease;
    `;
    const loader = document.createElement('globe-loader');
    loader.setAttribute('size', opts.size || 200);
    _overlay.appendChild(loader);
    document.body.appendChild(_overlay);
    requestAnimationFrame(() => { _overlay.style.opacity = '1'; });
  };
  GlobeLoader.hideOverlay = function () {
    if (!_overlay) return;
    const o = _overlay; _overlay = null;
    o.style.opacity = '0';
    setTimeout(() => o.remove(), 220);
  };

  window.GlobeLoader = GlobeLoader;
})();
