// Pictures with Word's Bildformat settings: Korrekturen (sharpen, brightness, contrast), Farbe
// (saturation, colour temperature, recolouring), Transparenz, cropping, rotation, the picture
// border, effects (shadow, reflection, glow, soft edges, bevel) and the frame shape of the
// picture styles. Drawn with CSS and SVG filters; the values are those Word writes in the
// picture's DrawingML (read and written by the ooxml package).

export interface PictureShadow {
  readonly kind: "outer" | "inner";
  readonly color: string;
  /** 0 (opaque) to 1 (invisible). */
  readonly transparency: number;
  /** Blur radius and distance in points; direction in degrees (0 = right, 90 = down). */
  readonly blur: number;
  readonly dist: number;
  readonly dir: number;
  /** Size of an outer shadow relative to the picture (perspective shadows are scaled and skewed). */
  readonly sx?: number;
  readonly sy?: number;
  /** Skew in degrees and alignment (DrawingML kx, ky, algn), kept for writing. */
  readonly kx?: number;
  readonly ky?: number;
  readonly algn?: string;
  readonly rotWithShape?: boolean;
}

/** A fill (shapes; the background of pictures): solid, gradient or none. */
export type Fill =
  | { readonly kind: "none" }
  | { readonly kind: "solid"; readonly color: string; readonly transparency?: number }
  | {
    readonly kind: "gradient"; readonly stops: readonly { readonly pos: number; readonly color: string; readonly transparency?: number }[];
    /** Linear: angle in degrees (0 = left to right, 90 = top to bottom); path: "circle", "rect" or "shape". */
    readonly angle?: number; readonly path?: string;
  }
  /** Picture, texture or pattern fills: drawn in their main colour. */
  | { readonly kind: "other"; readonly color: string };

/** An arrowhead at the start (head) or end (tail) of a line (a:headEnd, a:tailEnd). */
export interface LineEnd {
  readonly type: "triangle" | "stealth" | "diamond" | "oval" | "arrow";
  readonly w?: "sm" | "med" | "lg";
  readonly len?: "sm" | "med" | "lg";
}

export interface PictureFormat {
  /** Helligkeit and Kontrast, −1…1 (0 = unchanged). */
  readonly brightness?: number;
  readonly contrast?: number;
  /** Schärfen (positive) or Weichzeichnen (negative), −1…1. */
  readonly sharpen?: number;
  /** Farbsättigung, 0…4 (1 = unchanged). */
  readonly saturation?: number;
  /** Farbton as colour temperature in Kelvin (6500 = unchanged). */
  readonly temperature?: number;
  /** Neu einfärben. */
  readonly recolor?:
    | { readonly kind: "grayscale" | "sepia" | "washout" }
    | { readonly kind: "bw"; readonly threshold: number }
    | { readonly kind: "duotone"; readonly dark: string; readonly light: string };
  /** Transparenz, 0…1. */
  readonly transparency?: number;
  /** Transparente Farbe bestimmen: this colour (hex) is see-through (a:clrChange). */
  readonly transparentColor?: string;
  /** Cropped share of the original picture on each side (0…1, may be negative: padding). */
  readonly crop?: { readonly l: number; readonly t: number; readonly r: number; readonly b: number };
  /** Drehen: degrees clockwise; Horizontal/Vertikal kippen. */
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /** Frame shape (DrawingML preset: "rect", "roundRect", "ellipse", "snip2DiagRect" …). */
  readonly geometry?: string;
  /** The shape's adjust values (a:avLst content, e.g. the corner radius), as XML. */
  readonly geometryAdj?: string;
  /** Bildrahmen; `null` = explicitly none. */
  readonly line?: {
    readonly color: string; readonly width: number; readonly dash?: string; readonly compound?: string;
    /** Line ends ("sq", "rnd") and mitred corners, as Word's picture styles set them. */
    readonly cap?: string; readonly miter?: boolean;
    /**
     * Shapes: line transparency, corner join ("round", "bevel", "miter"), arrowheads. A colour
     * "" or a width of -1 means the shape style's (as Word writes arrow lines: only the ends).
     */
    readonly transparency?: number; readonly join?: string;
    readonly head?: LineEnd; readonly tail?: LineEnd;
  } | null;
  /** Shapes: the fill as read (fillXml keeps it for writing back unchanged). */
  readonly fill?: Fill;
  /** Older brightness/contrast (a:lum), which Word uses for "Ausgewaschen". */
  readonly lum?: { readonly bright: number; readonly contrast: number };
  /** Fill behind the picture, 3D scene and shape (a:solidFill, a:scene3d, a:sp3d), as XML. */
  readonly fillXml?: string;
  readonly scene3dXml?: string;
  readonly sp3dXml?: string;
  /** Space the effects take around the picture (wp:effectExtent), in points. */
  readonly extent?: { readonly l: number; readonly t: number; readonly r: number; readonly b: number };
  /** Bildformatvorlage applied last (index in Word's gallery), for the gallery's frame. */
  readonly preset?: number;
  readonly shadow?: PictureShadow;
  /** Spiegelung: size share of the picture shown, distance in points, start transparency. */
  readonly reflection?: { readonly size: number; readonly dist: number; readonly transparency: number; readonly blur: number; readonly endA?: number };
  /** Leuchten: radius in points. */
  readonly glow?: { readonly color: string; readonly radius: number; readonly transparency: number };
  /** Weiche Kanten: radius in points. */
  readonly softEdge?: number;
  /** Abschrägung (top bevel) width and height in points. */
  readonly bevel?: { readonly kind: string; readonly w: number; readonly h: number };
}

export const NO_PICTURE_FORMAT: PictureFormat = {};

/** DrawingML camera presets: latitude, longitude, revolution in degrees (approximations). */
const CAMERA: Record<string, [number, number, number]> = {
  isometricOffAxis1Left: [10, 45, 0], isometricOffAxis1Right: [10, -45, 0], isometricOffAxis2Left: [10, 45, 0], isometricOffAxis2Right: [10, -45, 0],
  isometricOffAxis1Top: [45, 0, 0], isometricOffAxis2Top: [45, 0, 0], isometricTopUp: [45, 0, 45], isometricTopDown: [45, 0, -45],
  isometricLeftDown: [0, 45, 0], isometricRightUp: [0, -45, 0], obliqueTopLeft: [10, 10, 0], obliqueTopRight: [10, -10, 0],
  obliqueBottomLeft: [-10, 10, 0], obliqueBottomRight: [-10, -10, 0], perspectiveLeft: [0, 20, 0], perspectiveRight: [0, -20, 0],
  perspectiveAbove: [20, 0, 0], perspectiveBelow: [-20, 0, 0], perspectiveAboveLeftFacing: [15, 20, 0], perspectiveAboveRightFacing: [15, -20, 0],
  perspectiveContrastingLeftFacing: [5, 25, 0], perspectiveContrastingRightFacing: [5, -25, 0], perspectiveHeroicLeftFacing: [-10, 25, 0],
  perspectiveHeroicRightFacing: [-10, -25, 0], perspectiveHeroicExtremeLeftFacing: [-15, 35, 0], perspectiveHeroicExtremeRightFacing: [-15, -35, 0],
  perspectiveRelaxed: [25, 0, 0], perspectiveRelaxedModerately: [15, 0, 0],
};

/** True when nothing about the picture's look is changed. */
export function plainPicture(f: PictureFormat | undefined): boolean {
  return !f || Object.values(f).every((v) => v === undefined);
}

const PT = 96 / 72;
let defs: SVGSVGElement | null = null;
const filters = new Map<string, string>();

/** An SVG filter (cached by its parameters) for the colour settings; "" when none is needed. */
function colorFilter(f: PictureFormat): string {
  const parts: string[] = [];
  if (f.lum && (f.lum.bright || f.lum.contrast)) {
    // Older a:lum (Word's "Ausgewaschen"): scaled towards black, then lightened.
    const k = 1 + f.lum.contrast, i = f.lum.bright;
    parts.push(`<feComponentTransfer><feFuncR type="linear" slope="${k}" intercept="${i}"/><feFuncG type="linear" slope="${k}" intercept="${i}"/><feFuncB type="linear" slope="${k}" intercept="${i}"/></feComponentTransfer>`);
  }
  const b = f.brightness ?? 0, c = f.contrast ?? 0;
  if (b || c) {
    // Office: contrast scales around the middle grey, brightness shifts.
    const k = c >= 0 ? 1 / Math.max(0.01, 1 - c) : 1 + c;
    const i = 0.5 - 0.5 * k + b;
    parts.push(`<feComponentTransfer><feFuncR type="linear" slope="${k}" intercept="${i}"/><feFuncG type="linear" slope="${k}" intercept="${i}"/><feFuncB type="linear" slope="${k}" intercept="${i}"/></feComponentTransfer>`);
  }
  if (f.saturation !== undefined && f.saturation !== 1) parts.push(`<feColorMatrix type="saturate" values="${f.saturation}"/>`);
  if (f.temperature !== undefined && f.temperature !== 6500) {
    // Warmer below 6500 K (more red, less blue), cooler above.
    const t = Math.max(-1, Math.min(1, (6500 - f.temperature) / 4500));
    const r = 1 + 0.18 * t, bl = 1 - 0.25 * t;
    parts.push(`<feColorMatrix type="matrix" values="${r} 0 0 0 0  0 1 0 0 0  0 0 ${bl} 0 0  0 0 0 1 0"/>`);
  }
  const rc = f.recolor;
  if (rc) {
    if (rc.kind === "grayscale") parts.push('<feColorMatrix type="saturate" values="0"/>');
    else if (rc.kind === "washout") parts.push('<feComponentTransfer><feFuncR type="linear" slope="0.3" intercept="0.7"/><feFuncG type="linear" slope="0.3" intercept="0.7"/><feFuncB type="linear" slope="0.3" intercept="0.7"/></feComponentTransfer>');
    // Sepia as Word does it: black to a light tan (D9C3A5 tinted 50 %, saturated 180 %).
    else if (rc.kind === "sepia") parts.push('<feColorMatrix type="saturate" values="0"/>', '<feComponentTransfer><feFuncR type="table" tableValues="0 0.93"/><feFuncG type="table" tableValues="0 0.84"/><feFuncB type="table" tableValues="0 0.72"/></feComponentTransfer>');
    else if (rc.kind === "bw") {
      const steps = Array.from({ length: 100 }, (_, i) => (i / 100 >= rc.threshold ? 1 : 0)).join(" ");
      parts.push('<feColorMatrix type="saturate" values="0"/>', `<feComponentTransfer><feFuncR type="discrete" tableValues="${steps}"/><feFuncG type="discrete" tableValues="${steps}"/><feFuncB type="discrete" tableValues="${steps}"/></feComponentTransfer>`);
    } else if (rc.kind === "duotone") {
      const hex = (x: string) => [0, 2, 4].map((i) => parseInt(x.slice(i, i + 2), 16) / 255);
      const [d, l] = [hex(rc.dark), hex(rc.light)];
      parts.push('<feColorMatrix type="saturate" values="0"/>',
        `<feComponentTransfer><feFuncR type="table" tableValues="${d[0]} ${l[0]}"/><feFuncG type="table" tableValues="${d[1]} ${l[1]}"/><feFuncB type="table" tableValues="${d[2]} ${l[2]}"/></feComponentTransfer>`);
    }
  }
  const s = f.sharpen ?? 0;
  if (s > 0) {
    const a = s * 2;
    parts.push(`<feConvolveMatrix order="3" preserveAlpha="true" kernelMatrix="0 ${-a} 0 ${-a} ${1 + 4 * a} ${-a} 0 ${-a} 0"/>`);
  } else if (s < 0) parts.push(`<feGaussianBlur stdDeviation="${-s * 2.5}"/>`);
  if (!parts.length) return "";
  const key = parts.join("");
  let id = filters.get(key);
  if (!id) {
    id = `raabe-pic-${filters.size + 1}`;
    filters.set(key, id);
    if (!defs) {
      defs = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      defs.setAttribute("width", "0");
      defs.setAttribute("height", "0");
      defs.style.cssText = "position:absolute;width:0;height:0;overflow:hidden";
      defs.setAttribute("aria-hidden", "true");
      document.body.append(defs);
    }
    defs.insertAdjacentHTML("beforeend", `<filter id="${id}" color-interpolation-filters="sRGB">${key}</filter>`);
  }
  return `url(#${id})`;
}

const rgba = (hex: string, transparency: number) => {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) || 0);
  return `rgba(${r},${g},${b},${Math.max(0, Math.min(1, 1 - transparency))})`;
};

/** An adjust value of the shape ("adj" in 1/100000 of the shorter side). */
function adjust(f: PictureFormat | undefined, name: string, fallback: number): number {
  const m = new RegExp(`name="${name}" fmla="val (-?\\d+)"`).exec(f?.geometryAdj ?? "");
  return m ? Number(m[1]) / 100000 : fallback;
}

/** Corner shape of a frame (CSS border-radius / clip-path), for a w × h box in px. */
function frameShape(geometry: string | undefined, w: number, h: number, f?: PictureFormat): string {
  const m = Math.min(w, h);
  switch (geometry) {
    case "ellipse": return "border-radius:50%";
    case "roundRect": return `border-radius:${m * adjust(f, "adj", 0.1667)}px`;
    case "round2DiagRect": return `border-radius:${m * adjust(f, "adj1", 0.1667)}px 0`;
    case "snip2DiagRect": {
      const s = m * 0.1667;
      return `clip-path:polygon(0 0,${w - s}px 0,${w}px ${s}px,${w}px ${h}px,${s}px ${h}px,0 ${h - s}px)`;
    }
    case "snipRoundRect": {
      const s = m * 0.1667;
      return `clip-path:polygon(${s}px 0,${w - s}px 0,${w}px ${s}px,${w}px ${h}px,0 ${h}px,0 ${s}px)`;
    }
    default: return "";
  }
}

/**
 * A picture as Word draws it: `w` × `h` px is the frame (after cropping), `url` the original
 * picture. The element is positioned by the caller (left/top set on the returned element).
 */
export function pictureElement(url: string, w: number, h: number, f: PictureFormat | undefined): HTMLElement {
  const outer = document.createElement("div");
  outer.style.cssText = `position:absolute;width:${w}px;height:${h}px`;
  const frame = document.createElement("div");
  const shape = frameShape(f?.geometry, w, h, f);
  frame.style.cssText = `position:absolute;inset:0;overflow:hidden;${shape}`;
  const img = document.createElement("img");
  img.src = f?.transparentColor ? withTransparentColor(url, f.transparentColor, img) : url;
  img.draggable = false;
  const c = f?.crop;
  if (c && (c.l || c.t || c.r || c.b)) {
    const iw = w / Math.max(0.01, 1 - c.l - c.r), ih = h / Math.max(0.01, 1 - c.t - c.b);
    img.style.cssText = `position:absolute;left:${-c.l * iw}px;top:${-c.t * ih}px;width:${iw}px;height:${ih}px;max-width:none`;
  } else img.style.cssText = "position:absolute;left:0;top:0;width:100%;height:100%";
  if (f) {
    const filter = colorFilter(f);
    if (filter) img.style.filter = filter;
    if (f.transparency) img.style.opacity = String(1 - f.transparency);
  }
  frame.append(img);
  outer.append(frame);
  if (!f) return outer;
  // A fill behind the picture (Word's picture styles: light grey under transparent parts).
  const fill = /<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"(?:\/>|>(?:<a:(shade|tint) val="(\d+)"\/>)?)/.exec(f.fillXml ?? "");
  if (fill) {
    let [r, g, bl] = [0, 2, 4].map((i) => parseInt(fill[1].slice(i, i + 2), 16));
    const k = fill[3] ? Number(fill[3]) / 100000 : 1;
    if (fill[2] === "shade") [r, g, bl] = [r, g, bl].map((x) => x * k);
    else if (fill[2] === "tint") [r, g, bl] = [r, g, bl].map((x) => x + (255 - x) * (1 - k));
    const alpha = /<a:alpha val="(\d+)"\/>/.exec(f.fillXml ?? "");
    frame.style.background = `rgba(${Math.round(r)},${Math.round(g)},${Math.round(bl)},${alpha ? Number(alpha[1]) / 100000 : 1})`;
  }
  // Bildrahmen: centred on the frame's edge, as Word draws lines.
  if (f.line && f.line.width > 0) {
    const lw = f.line.width * PT;
    const ring = document.createElement("div");
    const dash = f.line.dash && f.line.dash !== "solid" ? (/dot/i.test(f.line.dash) && !/dash/i.test(f.line.dash) ? "dotted" : "dashed") : f.line.compound === "dbl" ? "double" : "solid";
    ring.style.cssText = `position:absolute;left:${-lw / 2}px;top:${-lw / 2}px;width:${w + lw}px;height:${h + lw}px;box-sizing:border-box;` +
      `border:${lw}px ${dash} #${f.line.color};opacity:${1 - (f.line.transparency ?? 0)};pointer-events:none;${frameShape(f.geometry, w + lw, h + lw, f)}`;
    outer.append(ring);
  }
  // Abschrägung: light on the top left, shade on the bottom right.
  if (f.bevel) {
    const bw = Math.max(1, f.bevel.w * PT);
    frame.style.boxShadow = `inset ${bw * 0.6}px ${bw * 0.6}px ${bw}px rgba(255,255,255,0.55), inset ${-bw * 0.6}px ${-bw * 0.6}px ${bw}px rgba(0,0,0,0.35)`;
    // The inner shadow sits on top of the picture.
    const top = document.createElement("div");
    top.style.cssText = `position:absolute;inset:0;pointer-events:none;${shape};box-shadow:${frame.style.boxShadow}`;
    outer.append(top);
  }
  if (f.softEdge && f.softEdge > 0) {
    const r = f.softEdge * PT;
    const mask = `linear-gradient(to right, transparent, #000 ${r}px, #000 calc(100% - ${r}px), transparent), linear-gradient(to bottom, transparent, #000 ${r}px, #000 calc(100% - ${r}px), transparent)`;
    frame.style.setProperty("mask-image", mask);
    frame.style.setProperty("-webkit-mask-image", mask);
    frame.style.setProperty("mask-composite", "intersect");
    frame.style.setProperty("-webkit-mask-composite", "source-in");
  }
  const effects: string[] = [];
  if (f.glow && f.glow.radius > 0) {
    const g = rgba(f.glow.color, f.glow.transparency);
    effects.push(`drop-shadow(0 0 ${f.glow.radius * PT * 0.5}px ${g})`, `drop-shadow(0 0 ${f.glow.radius * PT * 0.5}px ${g})`);
  }
  const sh = f.shadow;
  if (sh && sh.kind === "outer") {
    const rad = (sh.dir * Math.PI) / 180;
    effects.push(`drop-shadow(${Math.cos(rad) * sh.dist * PT}px ${Math.sin(rad) * sh.dist * PT}px ${sh.blur * PT * 0.5}px ${rgba(sh.color, sh.transparency)})`);
  } else if (sh && sh.kind === "inner") {
    const rad = (sh.dir * Math.PI) / 180;
    const inner = document.createElement("div");
    inner.style.cssText = `position:absolute;inset:0;pointer-events:none;${shape};` +
      `box-shadow:inset ${Math.cos(rad) * sh.dist * PT}px ${Math.sin(rad) * sh.dist * PT}px ${sh.blur * PT * 0.5}px ${rgba(sh.color, sh.transparency)}`;
    outer.append(inner);
  }
  if (effects.length) outer.style.filter = effects.join(" ");
  if (f.reflection && f.reflection.size > 0) {
    const start = 1 - f.reflection.size;
    const reflect = `below ${f.reflection.dist * PT}px linear-gradient(transparent ${start * 100}%, rgba(255,255,255,${(1 - f.reflection.transparency) * 0.9}))`;
    outer.style.setProperty("-webkit-box-reflect", reflect);
  }
  const tf: string[] = [];
  // 3D rotation of the scene's camera (3D-Drehung), approximated with CSS perspective.
  const cam = /<a:camera prst="([^"]+)"[^>]*?(?:\/>|>\s*<a:rot lat="(\d+)" lon="(\d+)" rev="(\d+)"\/>)/.exec(f.scene3dXml ?? "");
  if (cam) {
    const deg = (v: string | undefined) => { const d = Number(v ?? 0) / 60000; return d > 180 ? d - 360 : d; };
    const preset = CAMERA[cam[1]] ?? [0, 0, 0];
    const lat = deg(cam[2]) + preset[0], lon = deg(cam[3]) + preset[1], rev = deg(cam[4]) + preset[2];
    if (lat || lon || rev) tf.push(`perspective(${cam[1].startsWith("perspective") ? 600 : 4000}px)`, `rotateX(${-lat}deg)`, `rotateY(${lon}deg)`, `rotateZ(${rev}deg)`);
  }
  if (f.rotation) tf.push(`rotate(${f.rotation}deg)`);
  if (f.flipH) tf.push("scaleX(-1)");
  if (f.flipV) tf.push("scaleY(-1)");
  if (tf.length) outer.style.transform = tf.join(" ");
  return outer;
}

/**
 * A gallery preview: the picture with its effects (shadow, reflection, glow, frame) made small
 * enough that everything fits the tile, and clipped to it, so neighbouring tiles never overlap.
 */
export function previewElement(url: string, w: number, h: number, f: PictureFormat | undefined): HTMLElement {
  const box = document.createElement("div");
  box.style.cssText = `position:relative;width:${w}px;height:${h}px;overflow:hidden;flex:none`;
  // Drawn as a picture of a usual size (effects in proportion, as in Word's thumbnails), then
  // scaled down until it fits with its effects.
  const vw = 150, vh = (150 * h) / w;
  const ex = f ? effectExtent({ ...f, extent: undefined }, vh / PT) : { l: 0, t: 0, r: 0, b: 0 };
  const l = ex.l * PT, t = ex.t * PT, r = ex.r * PT, b = ex.b * PT;
  const fw = vw + l + r, fh = vh + t + b;
  const k = Math.min((w - 2) / fw, (h - 2) / fh);
  const inner = document.createElement("div");
  inner.style.cssText = `position:absolute;left:${(w - fw * k) / 2}px;top:${(h - fh * k) / 2}px;width:${fw}px;height:${fh}px;transform:scale(${k});transform-origin:0 0`;
  const el = pictureElement(url, vw, vh, f);
  el.style.left = `${l}px`;
  el.style.top = `${t}px`;
  inner.append(el);
  box.append(inner);
  return box;
}

const transparentCache = new Map<string, string>();

/**
 * The picture with one colour made see-through (Transparente Farbe bestimmen): the address of
 * a copy once it is made; until then the picture as it is, which `img` then swaps for it.
 */
function withTransparentColor(url: string, hex: string, img: HTMLImageElement): string {
  const key = `${url}|${hex}`;
  const done = transparentCache.get(key);
  if (done) return done;
  const src = new Image();
  src.onload = () => {
    const c = document.createElement("canvas");
    c.width = src.naturalWidth;
    c.height = src.naturalHeight;
    const g = c.getContext("2d");
    if (!g) return;
    g.drawImage(src, 0, 0);
    const data = g.getImageData(0, 0, c.width, c.height);
    const d = data.data;
    const r = parseInt(hex.slice(0, 2), 16), gg = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
    // Word changes exactly that colour; a step of tolerance keeps compressed pictures usable.
    for (let i = 0; i < d.length; i += 4) {
      if (Math.abs(d[i] - r) <= 2 && Math.abs(d[i + 1] - gg) <= 2 && Math.abs(d[i + 2] - b) <= 2) d[i + 3] = 0;
    }
    g.putImageData(data, 0, 0);
    c.toBlob((blob) => {
      if (!blob) return;
      const out = URL.createObjectURL(blob);
      transparentCache.set(key, out);
      if (img.src === url || img.getAttribute("src") === url) img.src = out;
    }, "image/png");
  };
  src.src = url;
  return url;
}

/** Space the effects take around the picture (points), when Word did not say. */
export function effectExtent(f: PictureFormat, h: number): { l: number; t: number; r: number; b: number } {
  if (f.extent) return f.extent;
  let l = 0, t = 0, r = 0, b = 0;
  const grow = (dl: number, dt: number, dr: number, db: number) => { l = Math.max(l, dl); t = Math.max(t, dt); r = Math.max(r, dr); b = Math.max(b, db); };
  if (f.line) grow(f.line.width / 2, f.line.width / 2, f.line.width / 2, f.line.width / 2);
  if (f.glow) grow(f.glow.radius, f.glow.radius, f.glow.radius, f.glow.radius);
  const s = f.shadow;
  if (s?.kind === "outer") {
    const dx = Math.cos((s.dir * Math.PI) / 180) * s.dist, dy = Math.sin((s.dir * Math.PI) / 180) * s.dist;
    grow(Math.max(0, s.blur - dx), Math.max(0, s.blur - dy), Math.max(0, s.blur + dx), Math.max(0, s.blur + dy));
  }
  if (f.reflection) grow(0, 0, 0, f.reflection.dist + h * f.reflection.size);
  return { l, t, r, b };
}
