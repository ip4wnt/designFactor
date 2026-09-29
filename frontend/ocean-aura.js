/*!
 * Ocean Aura — autonomous seamless liquid-contour loop (WebGL 1).
 *
 * Everything the author tunes lives in the two objects at the top:
 *   PALETTES — all artistically meaningful colours, by role.
 *   CONFIG   — loop duration, seed, active palette, field shape, performance,
 *              film grain (SVG feTurbulence overlay).
 *
 * Loop model
 *   phase = (elapsed % duration) / duration,  theta = 2π·phase.
 *   The field is 4D simplex noise sampled at (x, y, R·cosθ, R·sinθ):
 *   time moves along a closed circle in noise space at constant speed, so
 *   every value and every time-derivative (the noise kernel is C³) matches
 *   at the seam. Nothing is accumulated between frames; any frame is a pure
 *   function of (phase, seed, config, palette, size).
 *
 * 4D simplex noise: Ian McEwan, Stefan Gustavson (Ashima Arts),
 * "webgl-noise", MIT License — https://github.com/ashima/webgl-noise
 */
(function (root) {
  'use strict';

  /* ------------------------------------------------------------------ */
  /* PALETTES — change colours here only.                                */
  /* Roles, from the luminous seams inward to the deepest pockets:       */
  /*   ridgeCore  thin bright core of each seam                          */
  /*   ridgeGlow  soft halo around the core                              */
  /*   shallow    surface right next to a seam (dominant tone)           */
  /*   mid        surface part-way into a cell                           */
  /*   deep       cell interiors                                         */
  /*   deeper     large cell interiors (tint field high)                 */
  /*   deepest    rare saturated pockets in the largest cells            */
  /* Keep luminance falling from ridgeCore to deepest so the relief      */
  /* still reads after a palette swap.                                   */
  /* ------------------------------------------------------------------ */
  // Strict JSON: double quotes, no comments inside, no trailing commas —
  // the block between the braces can be copied to/from a .json file as is.
  // "reference" — sampled from the reference video; "glacier" — cold;
  // "ember" — warm; "dusk" — muted, dark blue-plum palette for use under
  // white UI: every colour keeps ≥ 4.5:1 contrast with #FFFFFF.
  var PALETTES = {
    "reference": {
      "ridgeCore": "#EAFEFF",
      "ridgeGlow": "#9CEEFE",
      "shallow":   "#06D2F8",
      "mid":       "#4CCBFE",
      "deep":      "#78B4FE",
      "deeper":    "#968CFE",
      "deepest":   "#8A2CF6"
    },
    "glacier": {
      "ridgeCore": "#F4FBFF",
      "ridgeGlow": "#BFE3F2",
      "shallow":   "#6FB6D3",
      "mid":       "#4C94B8",
      "deep":      "#34729A",
      "deeper":    "#26577E",
      "deepest":   "#1A3A5E"
    },
    "ember": {
      "ridgeCore": "#FFF8EC",
      "ridgeGlow": "#FFD9A8",
      "shallow":   "#FFB36B",
      "mid":       "#FF9460",
      "deep":      "#F2735E",
      "deeper":    "#D9587A",
      "deepest":   "#A8326E"
    },
    "abyss": {
      "ridgeCore": "#3AA7BA",
      "ridgeGlow": "#1B6680",
      "shallow":   "#0E3D5C",
      "mid":       "#0A2D4A",
      "deep":      "#08213C",
      "deeper":    "#061832",
      "deepest":   "#040E24"
    },
    "abyss-deep": {
      "ridgeCore": "#1B6A80",
      "ridgeGlow": "#0F455B",
      "shallow":   "#0A2B44",
      "mid":       "#08223A",
      "deep":      "#061A30",
      "deeper":    "#051329",
      "deepest":   "#030B1C"
    },
    "abyss-night": {
      "ridgeCore": "#124C5E",
      "ridgeGlow": "#0B3345",
      "shallow":   "#071F33",
      "mid":       "#06192B",
      "deep":      "#051425",
      "deeper":    "#040F1F",
      "deepest":   "#030916"
    },
    "abyss-soft": {
      "ridgeCore": "#25849A",
      "ridgeGlow": "#154F68",
      "shallow":   "#0C3450",
      "mid":       "#092842",
      "deep":      "#071E37",
      "deeper":    "#05162E",
      "deepest":   "#040D20"
    },
    "dusk": {
      "ridgeCore": "#767380",
      "ridgeGlow": "#596384",
      "shallow":   "#324063",
      "mid":       "#242F53",
      "deep":      "#152346",
      "deeper":    "#221C32",
      "deepest":   "#0A0E27"
    }
  };

  /* ------------------------------------------------------------------ */
  /* CONFIG — structure, timing and cost. No colours here.               */
  /* ------------------------------------------------------------------ */
  var CONFIG = {
    // Loop length in seconds. 175 s = period measured in the reference
    // (frame-difference minimum at 10 s ↔ 185 s ↔ 360 s, ±0.1 s).
    duration: 175,
    seed: 7,                 // fixed, deterministic; changes the pattern, not the timing
    palette: 'abyss-deep',         // a key of PALETTES, or a palette object with the same roles

    field: {
      cellsAcross: 1.92,     // noise frequency along the container's long side
                             // (4.2 matched the reference; /1.3, /1.2, /1.4 → cells ≈2.18× larger)
      warp: 0.42,            // domain-warp strength (organic, liquid contours)
      warpScale: 0.55,       // warp field frequency relative to the main field
      evolution: 1.6,        // radius of the time circle in noise units. Local tempo =
                             // 2π·evolution/duration; 1.6 @ 175 s matches the reference's
                             // frame-to-frame change (lag 0.5 s / 1 s / 2 s)
      ridgeCore: 0.013,     // seam core width (in field units; scaled with cellsAcross to keep on-screen thickness)
      ridgeGlow: 0.055,     // seam halo width (scaled likewise)
      depthRange: 0.55,      // |field| that maps to the deepest ramp stop
      tintScale: 0.45,       // frequency of the slow field that decides where violet pools
      tintEvolution: 2.0,    // radius of the tint field's own time circle
      tintAmount: 1.0,
      dither: 1.5            // static triangular (TPDF) dither in 1/255 steps: breaks the
                             // 8-bit banding in slow dark gradients; 0 = off
    },

    maxPixelRatio: 1.5,      // cap on devicePixelRatio
    renderScale: 1.0,        // extra multiplier (<1 renders fewer pixels, upscaled by CSS)

    // Film grain: an SVG <feTurbulence> overlay above the WebGL canvas.
    // Static and seeded (no per-frame randomness), so the loop stays
    // seamless; monochrome, so it never changes the palette's hues.
    grain: {
      enabled: true,
      opacity: 0.22,         // strength of the overlay
      frequency: 1.21,       // feTurbulence baseFrequency; higher = finer grain
                             // (0.85 / 0.7 → grain 30% smaller)
      octaves: 3,            // detail of the grain
      contrast: 1.8,         // slope applied to the noise before blending
      seed: 7,               // fixed seed of the grain pattern
      blend: 'overlay'       // CSS mix-blend-mode; overlay keeps exposure close
    }
  };

  var grainCounter = 0;
  var ROLES = ['ridgeCore', 'ridgeGlow', 'shallow', 'mid', 'deep', 'deeper', 'deepest'];

  var VERT = [
    'attribute vec2 aPos;',
    'void main(){ gl_Position = vec4(aPos, 0.0, 1.0); }'
  ].join('\n');

  var FRAG = [
    '#ifdef GL_FRAGMENT_PRECISION_HIGH',
    'precision highp float;',
    '#else',
    'precision mediump float;',
    '#endif',
    'uniform vec2  uRes;',
    'uniform vec2  uLoop;',      // (cosθ, sinθ)
    'uniform vec4  uSeed;',      // per-seed offsets in noise space
    'uniform float uCells;',
    'uniform float uWarp;',
    'uniform float uWarpScale;',
    'uniform float uEvo;',
    'uniform float uCore;',
    'uniform float uGlow;',
    'uniform float uDepth;',
    'uniform float uTintScale;',
    'uniform float uTintAmt;',
    'uniform float uTintEvo;',
    'uniform float uDither;',
    'uniform vec3  uCol[7];',
    '',
    'vec4 mod289(vec4 x){ return x - floor(x * (1.0/289.0)) * 289.0; }',
    'float mod289(float x){ return x - floor(x * (1.0/289.0)) * 289.0; }',
    'vec4 permute(vec4 x){ return mod289(((x*34.0)+10.0)*x); }',
    'float permute(float x){ return mod289(((x*34.0)+10.0)*x); }',
    'vec4 taylorInvSqrt(vec4 r){ return 1.79284291400159 - 0.85373472095314 * r; }',
    'float taylorInvSqrt(float r){ return 1.79284291400159 - 0.85373472095314 * r; }',
    'vec4 grad4(float j, vec4 ip){',
    '  const vec4 ones = vec4(1.0, 1.0, 1.0, -1.0);',
    '  vec4 p, s;',
    '  p.xyz = floor(fract(vec3(j) * ip.xyz) * 7.0) * ip.z - 1.0;',
    '  p.w = 1.5 - dot(abs(p.xyz), ones.xyz);',
    '  s = vec4(lessThan(p, vec4(0.0)));',
    '  p.xyz = p.xyz + (s.xyz*2.0 - 1.0) * s.www;',
    '  return p;',
    '}',
    'float snoise(vec4 v){',
    '  const vec4 C = vec4(0.138196601125011, 0.276393202250021, 0.414589803375032, -0.447213595499958);',
    '  vec4 i  = floor(v + dot(v, vec4(0.309016994374947451)));',
    '  vec4 x0 = v - i + dot(i, C.xxxx);',
    '  vec4 i0;',
    '  vec3 isX = step(x0.yzw, x0.xxx);',
    '  vec3 isYZ = step(x0.zww, x0.yyz);',
    '  i0.x = isX.x + isX.y + isX.z;',
    '  i0.yzw = 1.0 - isX;',
    '  i0.y += isYZ.x + isYZ.y;',
    '  i0.zw += 1.0 - isYZ.xy;',
    '  i0.z += isYZ.z;',
    '  i0.w += 1.0 - isYZ.z;',
    '  vec4 i3 = clamp(i0, 0.0, 1.0);',
    '  vec4 i2 = clamp(i0 - 1.0, 0.0, 1.0);',
    '  vec4 i1 = clamp(i0 - 2.0, 0.0, 1.0);',
    '  vec4 x1 = x0 - i1 + C.xxxx;',
    '  vec4 x2 = x0 - i2 + C.yyyy;',
    '  vec4 x3 = x0 - i3 + C.zzzz;',
    '  vec4 x4 = x0 + C.wwww;',
    '  i = mod289(i);',
    '  float j0 = permute(permute(permute(permute(i.w) + i.z) + i.y) + i.x);',
    '  vec4 j1 = permute(permute(permute(permute(',
    '             i.w + vec4(i1.w, i2.w, i3.w, 1.0))',
    '           + i.z + vec4(i1.z, i2.z, i3.z, 1.0))',
    '           + i.y + vec4(i1.y, i2.y, i3.y, 1.0))',
    '           + i.x + vec4(i1.x, i2.x, i3.x, 1.0));',
    '  vec4 ip = vec4(1.0/294.0, 1.0/49.0, 1.0/7.0, 0.0);',
    '  vec4 p0 = grad4(j0, ip);',
    '  vec4 p1 = grad4(j1.x, ip);',
    '  vec4 p2 = grad4(j1.y, ip);',
    '  vec4 p3 = grad4(j1.z, ip);',
    '  vec4 p4 = grad4(j1.w, ip);',
    '  vec4 norm = taylorInvSqrt(vec4(dot(p0,p0), dot(p1,p1), dot(p2,p2), dot(p3,p3)));',
    '  p0 *= norm.x; p1 *= norm.y; p2 *= norm.z; p3 *= norm.w;',
    '  p4 *= taylorInvSqrt(dot(p4,p4));',
    '  vec3 m0 = max(0.6 - vec3(dot(x0,x0), dot(x1,x1), dot(x2,x2)), 0.0);',
    '  vec2 m1 = max(0.6 - vec2(dot(x3,x3), dot(x4,x4)), 0.0);',
    '  m0 = m0 * m0; m1 = m1 * m1;',
    '  return 49.0 * (dot(m0*m0, vec3(dot(p0,x0), dot(p1,x1), dot(p2,x2)))',
    '               + dot(m1*m1, vec2(dot(p3,x3), dot(p4,x4))));',
    '}',
    '',
    'void main(){',
    // long-side normalised coordinates, centred: framing adapts to any aspect
    '  vec2 p = (gl_FragCoord.xy - 0.5 * uRes) / max(uRes.x, uRes.y) * uCells;',
    '  vec2 c = uLoop * uEvo;',                          // point on the time circle
    // domain warp (its own, slower circle keeps it periodic too)
    '  vec2 q = p * uWarpScale;',
    '  vec2 cw = c * 0.6;',
    '  vec2 w = vec2(snoise(vec4(q + uSeed.xy, cw)),',
    '                snoise(vec4(q + uSeed.zw, cw + 31.7)));',
    '  float h = snoise(vec4(p + uWarp * w, c + uSeed.yx));',
    '  float d = abs(h);',                                // distance-like: 0 on a seam
    // slow tint field: where interiors lean lavender/violet
    '  float m = snoise(vec4(p * uTintScale + uSeed.wz, uLoop * uTintEvo - 17.3));',
    '  float tint = clamp((m * 0.5 + 0.5) * uTintAmt, 0.0, 1.0);',
    '  float t = clamp(d / uDepth, 0.0, 1.0);',
    // ramp from seam outward
    '  vec3 interior = mix(uCol[3], mix(uCol[4], uCol[5], smoothstep(0.45, 0.85, tint)), smoothstep(0.1, 0.55, tint));',
    '  vec3 col = mix(uCol[2], interior, smoothstep(0.18, 0.95, t));',
    '  float pool = smoothstep(0.58, 0.9, tint) * smoothstep(0.4, 1.0, t);',
    '  col = mix(col, uCol[6], pool);',
    // luminous seam: halo then core
    '  float glow = exp(-(d*d) / (uGlow*uGlow));',
    '  float core = exp(-(d*d) / (uCore*uCore));',
    '  col = mix(col, uCol[1], glow * 0.75);',
    '  col = mix(col, uCol[0], core);',
    // static TPDF dither (two decorrelated interleaved-gradient-noise samples,
    // fixed per pixel — no per-frame randomness): removes 8-bit banding
    '  vec2 fc = gl_FragCoord.xy;',
    '  float n1 = fract(52.9829189 * fract(dot(fc, vec2(0.06711056, 0.00583715))));',
    '  float n2 = fract(52.9829189 * fract(dot(fc + vec2(47.0, 17.0), vec2(0.00583715, 0.06711056))));',
    '  col += (n1 + n2 - 1.0) * (uDither / 255.0);',
    '  gl_FragColor = vec4(col, 1.0);',
    '}'
  ].join('\n');

  /* ------------------------------------------------------------------ */
  function hexToRgb(hex) {
    var h = String(hex).replace('#', '');
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    var n = parseInt(h, 16);
    return [((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255];
  }

  function resolvePalette(p) {
    var pal = typeof p === 'string' ? PALETTES[p] : p;
    if (!pal) throw new Error('OceanAura: unknown palette "' + p + '"');
    return pal;
  }

  // Deterministic seed → 4 offsets (mulberry32). Computed once, never per frame.
  function seedOffsets(seed) {
    var a = (seed >>> 0) || 1;
    function rnd() {
      a = (a + 0x6D2B79F5) >>> 0;
      var t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    }
    return [rnd() * 60 - 30, rnd() * 60 - 30, rnd() * 60 - 30, rnd() * 60 - 30];
  }

  function compile(gl, type, src) {
    var s = gl.createShader(type);
    gl.shaderSource(s, src);
    gl.compileShader(s);
    if (!gl.getShaderParameter(s, gl.COMPILE_STATUS)) {
      var log = gl.getShaderInfoLog(s);
      gl.deleteShader(s);
      throw new Error('OceanAura shader: ' + log);
    }
    return s;
  }

  /* ------------------------------------------------------------------ */
  /* create(host, overrides?) → instance                                 */
  /* The host element is filled by a canvas that ignores all input.      */
  /* ------------------------------------------------------------------ */
  function create(host, overrides) {
    overrides = overrides || {};
    var cfg = {
      duration: overrides.duration != null ? overrides.duration : CONFIG.duration,
      seed: overrides.seed != null ? overrides.seed : CONFIG.seed,
      palette: overrides.palette != null ? overrides.palette : CONFIG.palette,
      field: Object.assign({}, CONFIG.field, overrides.field || {}),
      maxPixelRatio: overrides.maxPixelRatio != null ? overrides.maxPixelRatio : CONFIG.maxPixelRatio,
      renderScale: overrides.renderScale != null ? overrides.renderScale : CONFIG.renderScale,
      grain: Object.assign({}, CONFIG.grain, overrides.grain || {}),
      autoplay: overrides.autoplay !== false
    };

    if (host.__oceanAura) host.__oceanAura.destroy();   // never two loops on one host

    var canvas = document.createElement('canvas');
    canvas.setAttribute('aria-hidden', 'true');
    canvas.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;display:block;pointer-events:none;touch-action:none;user-select:none;';
    if (getComputedStyle(host).position === 'static') host.style.position = 'relative';
    host.style.overflow = 'hidden';
    host.style.pointerEvents = 'none';
    host.appendChild(canvas);

    /* film grain overlay (SVG feTurbulence) */
    var SVGNS = 'http://www.w3.org/2000/svg';
    var grainId = 'oa-grain-' + (++grainCounter);
    var grainSvg = document.createElementNS(SVGNS, 'svg');
    grainSvg.setAttribute('aria-hidden', 'true');
    grainSvg.setAttribute('focusable', 'false');
    grainSvg.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;display:block;pointer-events:none;';
    var gFilter = document.createElementNS(SVGNS, 'filter');
    gFilter.setAttribute('id', grainId);
    gFilter.setAttribute('x', '0'); gFilter.setAttribute('y', '0');
    gFilter.setAttribute('width', '100%'); gFilter.setAttribute('height', '100%');
    gFilter.setAttribute('color-interpolation-filters', 'sRGB');
    var gTurb = document.createElementNS(SVGNS, 'feTurbulence');
    gTurb.setAttribute('type', 'fractalNoise');
    gTurb.setAttribute('stitchTiles', 'stitch');
    var gMono = document.createElementNS(SVGNS, 'feColorMatrix');
    gMono.setAttribute('type', 'saturate'); gMono.setAttribute('values', '0');
    var gCurve = document.createElementNS(SVGNS, 'feComponentTransfer');
    var gFuncs = ['feFuncR', 'feFuncG', 'feFuncB'].map(function (n) {
      var f = document.createElementNS(SVGNS, n); f.setAttribute('type', 'linear'); gCurve.appendChild(f); return f;
    });
    var gAlpha = document.createElementNS(SVGNS, 'feFuncA');
    gAlpha.setAttribute('type', 'linear'); gAlpha.setAttribute('slope', '0'); gAlpha.setAttribute('intercept', '1');
    gCurve.appendChild(gAlpha);
    gFilter.appendChild(gTurb); gFilter.appendChild(gMono); gFilter.appendChild(gCurve);
    var gRect = document.createElementNS(SVGNS, 'rect');
    gRect.setAttribute('width', '100%'); gRect.setAttribute('height', '100%');
    gRect.setAttribute('filter', 'url(#' + grainId + ')');
    grainSvg.appendChild(gFilter); grainSvg.appendChild(gRect);
    host.appendChild(grainSvg);
    function applyGrain() {
      var g = cfg.grain;
      grainSvg.style.display = g.enabled ? 'block' : 'none';
      grainSvg.style.opacity = String(g.opacity);
      grainSvg.style.mixBlendMode = g.blend;
      gTurb.setAttribute('baseFrequency', String(g.frequency));
      gTurb.setAttribute('numOctaves', String(g.octaves));
      gTurb.setAttribute('seed', String(g.seed));
      var slope = g.contrast, icpt = 0.5 - 0.5 * g.contrast;   // contrast around mid-grey
      gFuncs.forEach(function (f) { f.setAttribute('slope', String(slope)); f.setAttribute('intercept', String(icpt)); });
    }
    applyGrain();

    var gl = canvas.getContext('webgl', { antialias: false, alpha: false, depth: false, stencil: false,
      premultipliedAlpha: false, preserveDrawingBuffer: !!overrides.preserveDrawingBuffer, powerPreference: 'default' });
    if (!gl) {
      host.style.background = resolvePalette(cfg.palette).shallow;
      return { destroy: function () { if (canvas.parentNode) canvas.parentNode.removeChild(canvas); if (grainSvg.parentNode) grainSvg.parentNode.removeChild(grainSvg); }, error: 'webgl-unavailable' };
    }

    var prog = gl.createProgram();
    var vs = compile(gl, gl.VERTEX_SHADER, VERT);
    var fs = compile(gl, gl.FRAGMENT_SHADER, FRAG);
    gl.attachShader(prog, vs);
    gl.attachShader(prog, fs);
    gl.linkProgram(prog);
    if (!gl.getProgramParameter(prog, gl.LINK_STATUS)) throw new Error('OceanAura link: ' + gl.getProgramInfoLog(prog));
    gl.useProgram(prog);

    var buf = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, buf);
    gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1, -1, 3, -1, -1, 3]), gl.STATIC_DRAW);
    var aPos = gl.getAttribLocation(prog, 'aPos');
    gl.enableVertexAttribArray(aPos);
    gl.vertexAttribPointer(aPos, 2, gl.FLOAT, false, 0, 0);

    var U = {};
    ['uRes', 'uLoop', 'uSeed', 'uCells', 'uWarp', 'uWarpScale', 'uEvo', 'uCore', 'uGlow',
      'uDepth', 'uTintScale', 'uTintAmt', 'uTintEvo', 'uDither', 'uCol'].forEach(function (n) { U[n] = gl.getUniformLocation(prog, n); });

    var colours = new Float32Array(21);    // reused; written only when the palette changes
    function applyPalette(p) {
      var pal = resolvePalette(p);
      for (var i = 0; i < ROLES.length; i++) {
        var rgb = hexToRgb(pal[ROLES[i]]);
        colours[i * 3] = rgb[0]; colours[i * 3 + 1] = rgb[1]; colours[i * 3 + 2] = rgb[2];
      }
      gl.uniform3fv(U.uCol, colours);
      host.style.background = pal.shallow;   // shown only before the first frame / on context loss
      cfg.palette = p;
    }
    function applyField() {
      var f = cfg.field;
      gl.uniform1f(U.uCells, f.cellsAcross);
      gl.uniform1f(U.uWarp, f.warp);
      gl.uniform1f(U.uWarpScale, f.warpScale);
      gl.uniform1f(U.uEvo, f.evolution);
      gl.uniform1f(U.uCore, f.ridgeCore);
      gl.uniform1f(U.uGlow, f.ridgeGlow);
      gl.uniform1f(U.uDepth, f.depthRange);
      gl.uniform1f(U.uTintScale, f.tintScale);
      gl.uniform1f(U.uTintAmt, f.tintAmount);
      gl.uniform1f(U.uTintEvo, f.tintEvolution);
      gl.uniform1f(U.uDither, f.dither);
    }
    function applySeed() {
      var o = seedOffsets(cfg.seed);
      gl.uniform4f(U.uSeed, o[0], o[1], o[2], o[3]);
    }
    applyPalette(cfg.palette);
    applyField();
    applySeed();

    /* sizing: follows the container, never restarts the clock */
    var width = 0, height = 0, dirtySize = true;
    function measure() {
      var r = host.getBoundingClientRect();
      var dpr = Math.min(root.devicePixelRatio || 1, cfg.maxPixelRatio) * cfg.renderScale;
      var w = Math.max(1, Math.round(r.width * dpr));
      var h = Math.max(1, Math.round(r.height * dpr));
      if (w !== width || h !== height) {
        width = w; height = h;
        canvas.width = w; canvas.height = h;
        gl.viewport(0, 0, w, h);
        gl.uniform2f(U.uRes, w, h);
      }
      dirtySize = false;
    }
    var ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(function () { dirtySize = true; if (!running) draw(lastElapsed); }) : null;
    if (ro) ro.observe(host);
    function onWinResize() { dirtySize = true; }
    root.addEventListener('resize', onWinResize);

    /* clock: counts only while the page is visible, so a hidden tab
       resumes exactly where it left off (no jump, no burst). */
    var accumulated = 0, resumedAt = 0, running = false, raf = 0, lastElapsed = 0, destroyed = false;
    function nowSec() { return performance.now() / 1000; }
    function elapsed() { return running ? accumulated + (nowSec() - resumedAt) : accumulated; }

    function phaseOf(t) {
      var d = cfg.duration;
      return ((t % d) + d) % d / d;        // phase = (elapsed % duration) / duration
    }

    function draw(t) {
      if (dirtySize) measure();
      var th = 2 * Math.PI * phaseOf(t);
      gl.uniform2f(U.uLoop, Math.cos(th), Math.sin(th));
      gl.drawArrays(gl.TRIANGLES, 0, 3);
      lastElapsed = t;
    }

    function frame() {
      raf = 0;
      if (!running) return;
      draw(elapsed());
      raf = root.requestAnimationFrame(frame);
    }
    function start() {
      if (running || destroyed) return;
      running = true;
      resumedAt = nowSec();
      if (!raf) raf = root.requestAnimationFrame(frame);
    }
    function stop() {
      if (!running) return;
      accumulated = elapsed();
      running = false;
      if (raf) { root.cancelAnimationFrame(raf); raf = 0; }
    }
    function onVisibility() { if (document.hidden) stop(); else if (cfg.autoplay) start(); }
    document.addEventListener('visibilitychange', onVisibility);

    function onLost(e) { e.preventDefault(); stop(); }
    function onRestored() {
      // rebuild is simplest and rare; keep phase
      var keep = accumulated, pal = cfg.palette;
      inst.destroy();
      var fresh = create(host, Object.assign({}, overrides, { palette: pal }));
      fresh.seek(keep);
    }
    canvas.addEventListener('webglcontextlost', onLost, false);
    canvas.addEventListener('webglcontextrestored', onRestored, false);

    var inst = {
      canvas: canvas,
      config: cfg,
      /** Swap palette live: name from PALETTES or an object with the same roles. Phase untouched. */
      setPalette: function (p) { applyPalette(p); if (!running) draw(lastElapsed); },
      /** Change loop length; phase is preserved (elapsed is rescaled). */
      setDuration: function (d) {
        if (!(d > 0)) return;
        var ph = phaseOf(elapsed());
        var wasRunning = running; stop();
        cfg.duration = d; accumulated = ph * d;
        if (wasRunning) start(); else draw(accumulated);
      },
      setField: function (f) { Object.assign(cfg.field, f); applyField(); if (!running) draw(lastElapsed); },
      setSeed: function (s) { cfg.seed = s; applySeed(); if (!running) draw(lastElapsed); },
      /** Update film grain live, e.g. setGrain({ opacity: 0.15 }) or setGrain({ enabled: false }). */
      setGrain: function (g) { Object.assign(cfg.grain, g); applyGrain(); },
      phase: function () { return phaseOf(elapsed()); },
      elapsed: elapsed,
      /** Deterministic render at an arbitrary time (tests, stills). */
      renderAt: function (t) { draw(t); },
      seek: function (t) { var r = running; stop(); accumulated = t; if (r) start(); else draw(t); },
      play: start,
      pause: stop,
      destroy: function () {
        if (destroyed) return;
        destroyed = true;
        stop();
        document.removeEventListener('visibilitychange', onVisibility);
        root.removeEventListener('resize', onWinResize);
        if (ro) ro.disconnect();
        canvas.removeEventListener('webglcontextlost', onLost);
        canvas.removeEventListener('webglcontextrestored', onRestored);
        gl.deleteBuffer(buf); gl.deleteProgram(prog); gl.deleteShader(vs); gl.deleteShader(fs);
        var lose = gl.getExtension('WEBGL_lose_context'); if (lose) lose.loseContext();
        if (canvas.parentNode) canvas.parentNode.removeChild(canvas);
        if (grainSvg.parentNode) grainSvg.parentNode.removeChild(grainSvg);
        if (host.__oceanAura === inst) host.__oceanAura = null;
      }
    };
    host.__oceanAura = inst;

    measure();
    draw(0);
    if (cfg.autoplay && !document.hidden) start();
    return inst;
  }

  root.OceanAura = { CONFIG: CONFIG, PALETTES: PALETTES, create: create };
})(typeof window !== 'undefined' ? window : this);