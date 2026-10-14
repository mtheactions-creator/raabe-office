// Preset geometries (ECMA-376 presetShapeDefinitions) in the compact syntax described in
// geometry.ts. Checked against Word's PDF export with tools/feel-lab/check_shapes.py.

const ELLIPSE = "M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z";
const ELLIPSE_RECT = `
gd idx cos wd2 2700000
gd idy sin hd2 2700000
gd il +- hc 0 idx
gd ir +- hc idx 0
gd it +- vc 0 idy
gd ib +- vc idy 0
rect il it ir ib`;
const RECT = "M l t L r t L r b L l b Z";

const CLOUD = "M 3900 14370 A 6753 9190 -11429249 7426832 A 5333 7267 -8646143 5396714 A 4365 5945 -8748475 5983381 A 4857 6595 -7859164 7034504 A 5333 7273 -4722533 6541615 A 6775 9220 -2776035 7816140 A 5785 7867 37501 6842000 A 6752 9215 1347096 6910353 A 7720 10543 3974558 4542661 A 4360 5918 -16496525 8804134 A 4345 5945 -14809710 9151131 Z";
const CLOUD_ARCS = "M 4693 26177 A 4345 5945 5204520 1585770 M 6928 34899 A 4360 5918 4416628 686848 M 16478 39090 A 6752 9215 8257449 844866 M 28827 34751 A 6752 9215 387196 959901 M 34129 22954 A 5785 7867 -4217541 4255042 M 41798 15354 A 5333 7273 1819082 1665090 M 38324 5426 A 4857 6595 -824660 891534 M 29078 3952 A 4857 6595 -8950887 1091722 M 22141 4720 A 4365 5945 -9809656 1061181 M 14000 5192 A 6753 9190 -4002417 739161 M 4127 15789 A 6753 9190 9459261 711490";

/** A star with n points (16, 24, 32): outer points on the ellipse, inner ones at adj. */
function star(n: number, adj: number): string {
  const lines = [`av adj ${adj}`, "gd a pin 0 adj 50000", "gd iwd2 */ wd2 a 50000", "gd ihd2 */ hd2 a 50000"];
  const cmds: string[] = [];
  for (let k = 0; k < n; k++) {
    const ao = 10800000 + (k * 21600000) / n;
    const ai = ao + 10800000 / n;
    lines.push(`gd ox${k} cos wd2 ${ao}`, `gd oy${k} sin hd2 ${ao}`, `gd x${k} +- hc ox${k} 0`, `gd y${k} +- vc oy${k} 0`);
    lines.push(`gd ix${k} cos iwd2 ${ai}`, `gd iy${k} sin ihd2 ${ai}`, `gd sx${k} +- hc ix${k} 0`, `gd sy${k} +- vc iy${k} 0`);
    cmds.push(`${k ? "L" : "M"} x${k} y${k}`, `L sx${k} sy${k}`);
  }
  lines.push("gd idx cos iwd2 2700000", "gd idy sin ihd2 2700000", "gd il +- hc 0 idx", "gd ir +- hc idx 0", "gd it +- vc 0 idy", "gd ib +- vc idy 0", "gd yAdj +- vc 0 ihd2");
  lines.push("hxy - 0 0 adj 0 50000 hc yAdj", "rect il it ir ib", `path: ${cmds.join(" ")} Z`);
  return lines.join("\n");
}

/** The line callouts: callout1..3, accentCallout, borderCallout, accentBorderCallout. */
function callout(points: number, border: boolean, accent: boolean): string {
  const defaults = [[18750, -8333, 112500, -38333], [18750, -8333, 18750, -16667, 112500, -46667], [18750, -8333, 18750, -16667, 100000, -16667, 112963, -8333]][points - 1];
  const lines: string[] = [];
  defaults.forEach((v, i) => lines.push(`av adj${i + 1} ${v}`));
  const line: string[] = [];
  for (let i = 0; i < points + 1; i++) {
    lines.push(`gd y${i + 1} */ h adj${2 * i + 1} 100000`, `gd x${i + 1} */ w adj${2 * i + 2} 100000`);
    lines.push(`hxy adj${2 * i + 2} -2147483647 2147483647 adj${2 * i + 1} -2147483647 2147483647 x${i + 1} y${i + 1}`);
    line.push(`${i ? "L" : "M"} x${i + 1} y${i + 1}`);
  }
  lines.push("rect l t r b", `path${border ? "" : " stroke=false"}: ${RECT}`);
  if (accent) lines.push("path fill=none: M x1 t L x1 b");
  lines.push(`path fill=none: ${line.join(" ")}`);
  return lines.join("\n");
}

export const PRESET_DEFS: Readonly<Record<string, string>> = {
  // ----- lines and connectors ---------------------------------------------------------------
  line: `
path fill=none: M l t L r b`,
  straightConnector1: `
path fill=none: M l t L r b`,
  bentConnector2: `
path fill=none: M l t L r t L r b`,
  bentConnector3: `
av adj1 50000
gd x1 */ w adj1 100000
hxy adj1 -2147483647 2147483647 - 0 0 x1 vc
path fill=none: M l t L x1 t L x1 b L r b`,
  curvedConnector3: `
av adj1 50000
gd x2 */ w adj1 100000
gd x1 +/ l x2 2
gd x3 +/ r x2 2
gd y3 */ h 3 4
hxy adj1 -2147483647 2147483647 - 0 0 x2 vc
path fill=none: M l t C x1 t x2 hd4 x2 vc C x2 y3 x3 b r b`,

  // ----- rectangles -------------------------------------------------------------------------
  rect: `
path: ${RECT}`,
  roundRect: `
av adj 16667
gd a pin 0 adj 50000
gd x1 */ ss a 100000
gd x2 +- r 0 x1
gd y2 +- b 0 x1
gd il */ x1 29289 100000
gd ir +- r 0 il
gd ib +- b 0 il
hxy adj 0 50000 - 0 0 x1 t
rect il il ir ib
path: M l x1 A x1 x1 cd2 cd4 L x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4 L x1 b A x1 x1 cd4 cd4 Z`,
  snip1Rect: `
av adj 16667
gd a pin 0 adj 50000
gd dx1 */ ss a 100000
gd x1 +- r 0 dx1
gd it */ dx1 1 2
gd ir +/ x1 r 2
hxy adj 0 50000 - 0 0 x1 t
rect l it ir b
path: M l t L x1 t L r dx1 L r b L l b Z`,
  snip2SameRect: `
av adj1 16667
av adj2 0
gd a1 pin 0 adj1 50000
gd a2 pin 0 adj2 50000
gd tx1 */ ss a1 100000
gd tx2 +- r 0 tx1
gd bx1 */ ss a2 100000
gd bx2 +- r 0 bx1
gd by1 +- b 0 bx1
gd d +- tx1 0 bx1
gd dx ?: d tx1 bx1
gd il */ dx 1 2
gd ir +- r 0 il
gd it */ tx1 1 2
gd ib +/ by1 b 2
hxy adj1 0 50000 - 0 0 tx2 t
hxy adj2 0 50000 - 0 0 bx1 b
rect il it ir ib
path: M tx1 t L tx2 t L r tx1 L r by1 L bx2 b L bx1 b L l by1 L l tx1 Z`,
  snip2DiagRect: `
av adj1 0
av adj2 16667
gd a1 pin 0 adj1 50000
gd a2 pin 0 adj2 50000
gd lx1 */ ss a1 100000
gd lx2 +- r 0 lx1
gd ly1 +- b 0 lx1
gd rx1 */ ss a2 100000
gd rx2 +- r 0 rx1
gd ry1 +- b 0 rx1
gd d +- lx1 0 rx1
gd dx ?: d lx1 rx1
gd il */ dx 1 2
gd ir +- r 0 il
gd ib +- b 0 il
hxy adj1 0 50000 - 0 0 lx1 t
hxy adj2 0 50000 - 0 0 rx2 t
rect il il ir ib
path: M lx1 t L rx2 t L r rx1 L r ly1 L lx2 b L rx1 b L l ry1 L l lx1 Z`,
  snipRoundRect: `
av adj1 16667
av adj2 16667
gd a1 pin 0 adj1 50000
gd a2 pin 0 adj2 50000
gd x1 */ ss a1 100000
gd dx2 */ ss a2 100000
gd x2 +- r 0 dx2
gd il */ x1 29289 100000
gd ir +/ x2 r 2
hxy adj1 0 50000 - 0 0 x1 t
hxy adj2 0 50000 - 0 0 x2 t
rect il il ir b
path: M x1 t L x2 t L r dx2 L r b L l b L l x1 A x1 x1 cd2 cd4 Z`,
  round1Rect: `
av adj 16667
gd a pin 0 adj 50000
gd dx1 */ ss a 100000
gd x1 +- r 0 dx1
gd idx */ dx1 29289 100000
gd ir +- r 0 idx
hxy adj 0 50000 - 0 0 x1 t
rect l t ir b
path: M l t L x1 t A dx1 dx1 3cd4 cd4 L r b L l b Z`,
  round2SameRect: `
av adj1 16667
av adj2 0
gd a1 pin 0 adj1 50000
gd a2 pin 0 adj2 50000
gd tx1 */ ss a1 100000
gd tx2 +- r 0 tx1
gd bx1 */ ss a2 100000
gd bx2 +- r 0 bx1
gd by1 +- b 0 bx1
gd d +- tx1 0 bx1
gd tdx */ tx1 29289 100000
gd bdx */ bx1 29289 100000
gd il ?: d tdx bdx
gd ir +- r 0 il
gd ib +- b 0 bdx
hxy adj1 0 50000 - 0 0 tx2 t
hxy adj2 0 50000 - 0 0 bx1 b
rect il tdx ir ib
path: M tx1 t L tx2 t A tx1 tx1 3cd4 cd4 L r by1 A bx1 bx1 0 cd4 L bx1 b A bx1 bx1 cd4 cd4 L l tx1 A tx1 tx1 cd2 cd4 Z`,
  round2DiagRect: `
av adj1 16667
av adj2 0
gd a1 pin 0 adj1 50000
gd a2 pin 0 adj2 50000
gd x1 */ ss a1 100000
gd y1 +- b 0 x1
gd a */ ss a2 100000
gd x2 +- r 0 a
gd y2 +- b 0 a
gd dx1 */ x1 29289 100000
gd dx2 */ a 29289 100000
gd d +- dx1 0 dx2
gd dx ?: d dx1 dx2
gd ir +- r 0 dx
gd ib +- b 0 dx
hxy adj1 0 50000 - 0 0 x1 t
hxy adj2 0 50000 - 0 0 x2 t
rect dx dx ir ib
path: M x1 t L x2 t A a a 3cd4 cd4 L r y1 A x1 x1 0 cd4 L a b A a a cd4 cd4 L l x1 A x1 x1 cd2 cd4 Z`,

  // ----- basic shapes -----------------------------------------------------------------------
  ellipse: `${ELLIPSE_RECT}
path: ${ELLIPSE}`,
  triangle: `
av adj 50000
gd a pin 0 adj 100000
gd x1 */ w a 200000
gd x2 */ w a 100000
gd x3 +- x1 wd2 0
hxy adj 0 100000 - 0 0 x2 t
rect x1 vc x3 b
path: M l b L x2 t L r b Z`,
  rtTriangle: `
gd it */ h 7 12
gd ir */ w 7 12
gd ib */ h 11 12
rect l it ir ib
path: M l b L l t L r b Z`,
  parallelogram: `
av adj 25000
gd maxAdj */ 100000 w ss
gd a pin 0 adj maxAdj
gd x1 */ ss a 200000
gd x2 */ ss a 100000
gd x6 +- r 0 x2
gd x5 +- r 0 x1
gd q1 */ 5 a maxAdj
gd q2 +/ 1 q1 12
gd il */ q2 w 1
gd it */ q2 h 1
gd ir +- r 0 il
gd ib +- b 0 it
hxy adj 0 maxAdj - 0 0 x2 t
rect il it ir ib
path: M l b L x2 t L r t L x6 b Z`,
  trapezoid: `
av adj 25000
gd maxAdj */ 50000 w ss
gd a pin 0 adj maxAdj
gd x1 */ ss a 200000
gd x2 */ ss a 100000
gd x3 +- r 0 x2
gd x4 +- r 0 x1
gd il */ wd3 a maxAdj
gd it */ hd3 a maxAdj
gd ir +- r 0 il
hxy adj 0 maxAdj - 0 0 x2 t
rect il it ir b
path: M l b L x2 t L x3 t L r b Z`,
  diamond: `
gd ir */ w 3 4
gd ib */ h 3 4
rect wd4 hd4 ir ib
path: M l vc L hc t L r vc L hc b Z`,
  pentagon: `
av hf 105146
av vf 110557
gd swd2 */ wd2 hf 100000
gd shd2 */ hd2 vf 100000
gd svc */ vc vf 100000
gd dx1 cos swd2 1080000
gd dx2 cos swd2 18360000
gd dy1 sin shd2 1080000
gd dy2 sin shd2 18360000
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc dx2 0
gd x4 +- hc dx1 0
gd y1 +- svc 0 dy1
gd y2 +- svc 0 dy2
gd it */ y1 dx2 dx1
rect x2 it x3 y2
path: M x1 y1 L hc t L x4 y1 L x3 y2 L x2 y2 Z`,
  hexagon: `
av adj 25000
av vf 115470
gd maxAdj */ 50000 w ss
gd a pin 0 adj maxAdj
gd shd2 */ hd2 vf 100000
gd x1 */ ss a 100000
gd x2 +- r 0 x1
gd dy1 sin shd2 3600000
gd y1 +- vc 0 dy1
gd y2 +- vc dy1 0
gd q1 */ maxAdj -1 2
gd q2 +- a q1 0
gd q3 ?: q2 4 2
gd q4 ?: q2 3 2
gd q5 ?: q2 q1 0
gd q6 +/ a q5 q1
gd q7 */ q6 q4 -1
gd q8 +- q3 q7 0
gd il */ w q8 24
gd it */ h q8 24
gd ir +- r 0 il
gd ib +- b 0 it
hxy adj 0 maxAdj - 0 0 x1 t
rect il it ir ib
path: M l vc L x1 y1 L x2 y1 L r vc L x2 y2 L x1 y2 Z`,
  heptagon: `
av hf 102572
av vf 105210
gd swd2 */ wd2 hf 100000
gd shd2 */ hd2 vf 100000
gd svc */ vc vf 100000
gd dx1 */ swd2 97493 100000
gd dx2 */ swd2 78183 100000
gd dx3 */ swd2 43388 100000
gd dy1 */ shd2 62349 100000
gd dy2 */ shd2 22252 100000
gd dy3 */ shd2 90097 100000
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc 0 dx3
gd x4 +- hc dx3 0
gd x5 +- hc dx2 0
gd x6 +- hc dx1 0
gd y1 +- svc 0 dy1
gd y2 +- svc dy2 0
gd y3 +- svc dy3 0
gd ib +- b 0 y1
rect x2 y1 x5 ib
path: M x1 y2 L x2 y1 L hc t L x5 y1 L x6 y2 L x4 y3 L x3 y3 Z`,
  octagon: `
av adj 29289
gd a pin 0 adj 50000
gd x1 */ ss a 100000
gd x2 +- r 0 x1
gd y2 +- b 0 x1
gd il */ x1 1 2
gd ir +- r 0 il
gd ib +- b 0 il
hxy adj 0 50000 - 0 0 x1 t
rect il il ir ib
path: M l x1 L x1 t L x2 t L r x1 L r y2 L x2 b L x1 b L l y2 Z`,
  decagon: `
av vf 105146
gd shd2 */ hd2 vf 100000
gd dx1 cos wd2 2160000
gd dx2 cos wd2 4320000
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc dx2 0
gd x4 +- hc dx1 0
gd dy1 sin shd2 4320000
gd dy2 sin shd2 2160000
gd y1 +- vc 0 dy1
gd y2 +- vc 0 dy2
gd y3 +- vc dy2 0
gd y4 +- vc dy1 0
rect x1 y2 x4 y3
path: M l vc L x1 y2 L x2 y1 L x3 y1 L x4 y2 L r vc L x4 y3 L x3 y4 L x2 y4 L x1 y3 Z`,
  dodecagon: `
gd x1 */ w 2894 21600
gd x2 */ w 7906 21600
gd x3 */ w 13694 21600
gd x4 */ w 18706 21600
gd y1 */ h 2894 21600
gd y2 */ h 7906 21600
gd y3 */ h 13694 21600
gd y4 */ h 18706 21600
rect x1 y1 x4 y4
path: M l y2 L x1 y1 L x2 t L x3 t L x4 y1 L r y2 L r y3 L x4 y4 L x3 b L x2 b L x1 y4 L l y3 Z`,
  pie: `
av adj1 0
av adj2 16200000
gd stAng pin 0 adj1 21599999
gd enAng pin 0 adj2 21599999
gd sw1 +- enAng 0 stAng
gd sw2 +- sw1 21600000 0
gd swAng ?: sw1 sw1 sw2
gd wt1 sin wd2 stAng
gd ht1 cos hd2 stAng
gd dx1 cat2 wd2 ht1 wt1
gd dy1 sat2 hd2 ht1 wt1
gd x1 +- hc dx1 0
gd y1 +- vc dy1 0
gd wt2 sin wd2 enAng
gd ht2 cos hd2 enAng
gd dx2 cat2 wd2 ht2 wt2
gd dy2 sat2 hd2 ht2 wt2
gd x2 +- hc dx2 0
gd y2 +- vc dy2 0
gd idx cos wd2 2700000
gd idy sin hd2 2700000
gd il +- hc 0 idx
gd ir +- hc idx 0
gd it +- vc 0 idy
gd ib +- vc idy 0
hpolar - 0 0 adj1 0 21599999 x1 y1
hpolar - 0 0 adj2 0 21599999 x2 y2
rect il it ir ib
path: M x1 y1 A wd2 hd2 stAng swAng L hc vc Z`,
  chord: `
av adj1 2700000
av adj2 16200000
gd stAng pin 0 adj1 21599999
gd enAng pin 0 adj2 21599999
gd sw1 +- enAng 0 stAng
gd sw2 +- sw1 21600000 0
gd swAng ?: sw1 sw1 sw2
gd wt1 sin wd2 stAng
gd ht1 cos hd2 stAng
gd dx1 cat2 wd2 ht1 wt1
gd dy1 sat2 hd2 ht1 wt1
gd wt2 sin wd2 enAng
gd ht2 cos hd2 enAng
gd dx2 cat2 wd2 ht2 wt2
gd dy2 sat2 hd2 ht2 wt2
gd x1 +- hc dx1 0
gd y1 +- vc dy1 0
gd x2 +- hc dx2 0
gd y2 +- vc dy2 0
gd idx cos wd2 2700000
gd idy sin hd2 2700000
gd il +- hc 0 idx
gd ir +- hc idx 0
gd it +- vc 0 idy
gd ib +- vc idy 0
hpolar - 0 0 adj1 0 21599999 x1 y1
hpolar - 0 0 adj2 0 21599999 x2 y2
rect il it ir ib
path: M x1 y1 A wd2 hd2 stAng swAng Z`,
  teardrop: `
av adj 100000
gd a pin 0 adj 200000
gd r2 sqrt 2
gd tw */ r2 wd2 1
gd th */ r2 hd2 1
gd sw */ tw a 100000
gd sh */ th a 100000
gd dx1 cos sw 2700000
gd dy1 sin sh 2700000
gd x1 +- hc dx1 0
gd y1 +- vc 0 dy1
gd x2 +/ hc x1 2
gd y2 +/ vc y1 2
gd idx cos wd2 2700000
gd idy sin hd2 2700000
gd il +- hc 0 idx
gd ir +- hc idx 0
gd it +- vc 0 idy
gd ib +- vc idy 0
hxy adj 0 200000 - 0 0 x1 t
rect il it ir ib
path: M l vc A wd2 hd2 cd2 cd4 Q x2 t x1 y1 Q r y2 r vc A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z`,
  frame: `
av adj1 12500
gd a1 pin 0 adj1 50000
gd x1 */ ss a1 100000
gd x4 +- r 0 x1
gd y4 +- b 0 x1
hxy adj1 0 50000 - 0 0 x1 t
rect x1 x1 x4 y4
path: ${RECT} M x1 x1 L x1 y4 L x4 y4 L x4 x1 Z`,
  halfFrame: `
av adj1 33333
av adj2 33333
gd maxAdj2 */ 100000 w ss
gd a2 pin 0 adj2 maxAdj2
gd x1 */ ss a2 100000
gd g1 */ h x1 w
gd g2 +- h 0 g1
gd maxAdj1 */ 100000 g2 ss
gd a1 pin 0 adj1 maxAdj1
gd y1 */ ss a1 100000
gd dx2 */ y1 w h
gd x2 +- r 0 dx2
gd dy2 */ x1 h w
gd y2 +- b 0 dy2
hxy - 0 0 adj1 0 maxAdj1 l y1
hxy adj2 0 maxAdj2 - 0 0 x1 t
rect l t r b
path: M l t L r t L x2 y1 L x1 y1 L x1 y2 L l b Z`,
  corner: `
av adj1 50000
av adj2 50000
gd maxAdj1 */ 100000 h ss
gd maxAdj2 */ 100000 w ss
gd a1 pin 0 adj1 maxAdj1
gd a2 pin 0 adj2 maxAdj2
gd x1 */ ss a2 100000
gd dy1 */ ss a1 100000
gd y1 +- b 0 dy1
gd d +- w 0 h
gd it ?: d y1 t
gd ir ?: d r x1
hxy - 0 0 adj1 0 maxAdj1 l y1
hxy adj2 0 maxAdj2 - 0 0 x1 t
rect l it ir b
path: M l t L x1 t L x1 y1 L r y1 L r b L l b Z`,
  diagStripe: `
av adj 50000
gd a pin 0 adj 100000
gd x2 */ w a 100000
gd x1 */ x2 1 2
gd x3 +/ x2 r 2
gd y2 */ h a 100000
gd y1 */ y2 1 2
gd y3 +/ y2 b 2
hxy - 0 0 adj 0 100000 l y2
rect l t x3 y3
path: M l y2 L x2 t L r t L l b Z`,
  plus: `
av adj 25000
gd a pin 0 adj 50000
gd x1 */ ss a 100000
gd x2 +- r 0 x1
gd y2 +- b 0 x1
gd d +- w 0 h
gd il ?: d l x1
gd ir ?: d r x2
gd it ?: d x1 t
gd ib ?: d y2 b
hxy adj 0 50000 - 0 0 x1 t
rect il it ir ib
path: M l x1 L x1 x1 L x1 t L x2 t L x2 x1 L r x1 L r y2 L x2 y2 L x2 b L x1 b L x1 y2 L l y2 Z`,
  plaque: `
av adj 16667
gd a pin 0 adj 50000
gd x1 */ ss a 100000
gd x2 +- r 0 x1
gd y2 +- b 0 x1
gd il */ x1 70711 100000
gd ir +- r 0 il
gd ib +- b 0 il
hxy adj 0 50000 - 0 0 x1 t
rect il il ir ib
path: M l x1 A x1 x1 cd4 -5400000 L x2 t A x1 x1 cd2 -5400000 L r y2 A x1 x1 3cd4 -5400000 L x1 b A x1 x1 0 -5400000 Z`,
  can: `
av adj 25000
gd maxAdj */ 50000 h ss
gd a pin 0 adj maxAdj
gd y1 */ ss a 200000
gd y2 +- y1 y1 0
gd y3 +- b 0 y1
hxy - 0 0 adj 0 maxAdj hc y2
rect l y2 r y3
path stroke=false: M l y1 A wd2 y1 cd2 -10800000 L r y3 A wd2 y1 0 cd2 Z
path stroke=false fill=lighten: M l y1 A wd2 y1 cd2 cd2 A wd2 y1 0 cd2 Z
path fill=none: M r y1 A wd2 y1 0 cd2 A wd2 y1 cd2 cd2 L r y3 A wd2 y1 0 cd2 L l y1`,
  cube: `
av adj 25000
gd a pin 0 adj 100000
gd y1 */ ss a 100000
gd y4 +- b 0 y1
gd y2 */ y4 1 2
gd y3 +/ y1 b 2
gd x4 +- r 0 y1
gd x2 */ x4 1 2
gd x3 +/ y1 r 2
hxy - 0 0 adj 0 100000 l y1
rect l y1 x4 b
path stroke=false: M l y1 L x4 y1 L x4 b L l b Z
path stroke=false fill=darkenLess: M x4 y1 L r t L r y4 L x4 b Z
path stroke=false fill=lightenLess: M l y1 L y1 t L r t L x4 y1 Z
path fill=none: M l y1 L y1 t L r t L r y4 L x4 b L l b Z M l y1 L x4 y1 L r t M x4 y1 L x4 b`,
  bevel: `
av adj 12500
gd a pin 0 adj 50000
gd x1 */ ss a 100000
gd x2 +- r 0 x1
gd y2 +- b 0 x1
hxy adj 0 50000 - 0 0 x1 t
rect x1 x1 x2 y2
path stroke=false: M x1 x1 L x2 x1 L x2 y2 L x1 y2 Z
path stroke=false fill=lightenLess: M l t L r t L x2 x1 L x1 x1 Z
path stroke=false fill=darkenLess: M l b L x1 y2 L x2 y2 L r b Z
path stroke=false fill=lighten: M l t L x1 x1 L x1 y2 L l b Z
path stroke=false fill=darken: M r t L r b L x2 y2 L x2 x1 Z
path fill=none: ${RECT} M x1 x1 L x2 x1 L x2 y2 L x1 y2 Z M l t L x1 x1 M l b L x1 y2 M r t L x2 x1 M r b L x2 y2`,
  donut: `
av adj 25000
gd a pin 0 adj 50000
gd dr */ ss a 100000
gd iwd2 +- wd2 0 dr
gd ihd2 +- hd2 0 dr${ELLIPSE_RECT}
hxy adj 0 50000 - 0 0 dr vc
path: ${ELLIPSE} M dr vc A iwd2 ihd2 cd2 -5400000 A iwd2 ihd2 cd4 -5400000 A iwd2 ihd2 0 -5400000 A iwd2 ihd2 3cd4 -5400000 Z`,
  noSmoking: `
av adj 18750
gd a pin 0 adj 50000
gd dr */ ss a 100000
gd iwd2 +- wd2 0 dr
gd ihd2 +- hd2 0 dr
gd ang at2 w h
gd ct cos ihd2 ang
gd st sin iwd2 ang
gd m mod ct st 0
gd n */ iwd2 ihd2 m
gd drd2 */ dr 1 2
gd dang at2 n drd2
gd dang2 */ dang 2 1
gd swAng +- -10800000 dang2 0
gd t3 at2 w h
gd stAng1 +- t3 0 dang
gd stAng2 +- stAng1 0 cd2
gd ct1 cos ihd2 stAng1
gd st1 sin iwd2 stAng1
gd m1 mod ct1 st1 0
gd n1 */ iwd2 ihd2 m1
gd dx1 cos n1 stAng1
gd dy1 sin n1 stAng1
gd x1 +- hc dx1 0
gd y1 +- vc dy1 0
gd x2 +- hc 0 dx1
gd y2 +- vc 0 dy1${ELLIPSE_RECT}
hxy adj 0 50000 - 0 0 dr vc
path: ${ELLIPSE} M x1 y1 A iwd2 ihd2 stAng1 swAng Z M x2 y2 A iwd2 ihd2 stAng2 swAng Z`,
  blockArc: `
av adj1 10800000
av adj2 0
av adj3 25000
gd stAng pin 0 adj1 21599999
gd istAng pin 0 adj2 21599999
gd a3 pin 0 adj3 50000
gd sw11 +- istAng 0 stAng
gd sw12 +- sw11 21600000 0
gd swAng ?: sw11 sw11 sw12
gd iswAng +- 0 0 swAng
gd wt1 sin wd2 stAng
gd ht1 cos hd2 stAng
gd dx1 cat2 wd2 ht1 wt1
gd dy1 sat2 hd2 ht1 wt1
gd x1 +- hc dx1 0
gd y1 +- vc dy1 0
gd dr */ ss a3 100000
gd iwd2 +- wd2 0 dr
gd ihd2 +- hd2 0 dr
gd wt2 sin iwd2 istAng
gd ht2 cos ihd2 istAng
gd dx2 cat2 iwd2 ht2 wt2
gd dy2 sat2 ihd2 ht2 wt2
gd x2 +- hc dx2 0
gd y2 +- vc dy2 0
hpolar adj3 0 50000 adj1 0 21599999 x1 y1
rect l t r b
path: M x1 y1 A wd2 hd2 stAng swAng L x2 y2 A iwd2 ihd2 istAng iswAng Z`,
  foldedCorner: `
av adj 16667
gd a pin 0 adj 50000
gd dy2 */ ss a 100000
gd dy1 */ dy2 1 5
gd x1 +- r 0 dy2
gd x2 +- x1 dy1 0
gd y2 +- b 0 dy2
gd y1 +- y2 dy1 0
hxy adj 0 50000 - 0 0 x1 b
rect l t r y2
path stroke=false: M l t L r t L r y2 L x1 b L l b Z
path stroke=false fill=darkenLess: M x1 b L x2 y1 L r y2 Z
path fill=none: M x1 b L x2 y1 L r y2 L x1 b L l b L l t L r t L r y2`,
  smileyFace: `
av adj 4653
gd a pin -4653 adj 4653
gd x1 */ w 4969 21699
gd x2 */ w 6215 21600
gd x3 */ w 13135 21600
gd x4 */ w 16640 21600
gd y1 */ h 7570 21600
gd y3 */ h 16515 21600
gd dy2 */ h a 100000
gd y2 +- y3 0 dy2
gd y4 +- y3 dy2 0
gd dy3 */ h a 50000
gd y5 +- y4 dy3 0
gd wR */ w 1125 21600
gd hR */ h 1125 21600${ELLIPSE_RECT}
hxy - 0 0 adj -4653 4653 hc y4
path stroke=false: M l vc A wd2 hd2 cd2 21600000 Z
path stroke=false fill=darkenLess: M x2 y1 A wR hR cd2 21600000 M x3 y1 A wR hR cd2 21600000
path fill=none: M x1 y2 Q hc y5 x4 y2
path fill=none: M l vc A wd2 hd2 cd2 21600000 Z`,
  heart: `
gd dx1 */ w 49 48
gd dx2 */ w 10 48
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc dx2 0
gd x4 +- hc dx1 0
gd y1 +- t 0 hd3
gd il */ w 1 6
gd ir */ w 5 6
gd ib */ h 2 3
rect il hd4 ir ib
path: M hc hd4 C x3 y1 x4 hd4 hc b C x1 hd4 x2 y1 hc hd4 Z`,
  lightningBolt: `
gd x4 */ w 8757 21600
gd x9 */ w 13917 21600
gd y4 */ h 7437 21600
gd y10 */ h 14277 21600
rect x4 y4 x9 y10
path w=21600 h=21600: M 8472 0 L 12860 6080 L 11050 6797 L 16577 12007 L 14767 12877 L 21600 21600 L 10012 14915 L 12222 13987 L 5022 9705 L 7602 8382 L 0 3890 Z`,
  sun: `
av adj 25000
gd a pin 12500 adj 46875
gd g0 +- 50000 0 a
gd g1 */ g0 30274 32768
gd g2 */ g0 12540 32768
gd g3 +- g1 50000 0
gd g4 +- g2 50000 0
gd g5 +- 50000 0 g1
gd g6 +- 50000 0 g2
gd g7 */ g0 23170 32768
gd g8 +- 50000 g7 0
gd g9 +- 50000 0 g7
gd g10 */ g5 3 4
gd g11 */ g6 3 4
gd g12 +- g10 3662 0
gd g13 +- g11 3662 0
gd g14 +- g11 12500 0
gd g15 +- 100000 0 g10
gd g16 +- 100000 0 g12
gd g17 +- 100000 0 g13
gd g18 +- 100000 0 g14
gd ox1 */ w 18436 21600
gd oy1 */ h 3163 21600
gd ox2 */ w 3163 21600
gd oy2 */ h 18436 21600
gd x8 */ w g8 100000
gd x9 */ w g9 100000
gd x10 */ w g10 100000
gd x12 */ w g12 100000
gd x13 */ w g13 100000
gd x14 */ w g14 100000
gd x15 */ w g15 100000
gd x16 */ w g16 100000
gd x17 */ w g17 100000
gd x18 */ w g18 100000
gd x19 */ w a 100000
gd wR */ w g0 100000
gd hR */ h g0 100000
gd y8 */ h g8 100000
gd y9 */ h g9 100000
gd y10 */ h g10 100000
gd y12 */ h g12 100000
gd y13 */ h g13 100000
gd y14 */ h g14 100000
gd y15 */ h g15 100000
gd y16 */ h g16 100000
gd y17 */ h g17 100000
gd y18 */ h g18 100000
hxy adj 12500 46875 - 0 0 x19 vc
rect x9 y9 x8 y8
path: M r vc L x15 y18 L x15 y14 Z M ox1 oy1 L x16 y13 L x17 y12 Z M hc t L x18 y10 L x14 y10 Z M ox2 oy1 L x13 y12 L x12 y13 Z M l vc L x10 y14 L x10 y18 Z M ox2 oy2 L x12 y17 L x13 y16 Z M hc b L x14 y15 L x18 y15 Z M ox1 oy2 L x17 y16 L x16 y17 Z M x19 vc A wR hR cd2 21600000 Z`,
  moon: `
av adj 50000
gd a pin 0 adj 87500
gd g0 */ ss a 100000
gd g0w */ g0 w ss
gd g1 +- ss 0 g0
gd g2 */ g0 g0 g1
gd g3 */ ss ss g1
gd g4 */ g3 2 1
gd g5 +- g4 0 g2
gd g6 +- g5 0 g0
gd g6w */ g6 w ss
gd g7 */ g5 1 2
gd g8 +- g7 0 g0
gd dy1 */ g8 hd2 ss
gd g10h +- vc 0 dy1
gd g11h +- vc dy1 0
gd g12 */ g0 9598 32768
gd g12w */ g12 w ss
gd g13 +- ss 0 g12
gd q1 */ ss ss 1
gd q2 */ g13 g13 1
gd q3 +- q1 0 q2
gd q4 sqrt q3
gd dy4 */ q4 hd2 ss
gd g15h +- vc 0 dy4
gd g16h +- vc dy4 0
gd g17w +- g6w 0 g0w
gd g18w */ g17w 1 2
gd dx2p +- g0w g18w w
gd dx2 */ dx2p -1 1
gd dy2 */ hd2 -1 1
gd stAng1 at2 dx2 dy2
gd enAngp1 at2 dx2 hd2
gd enAng1 +- enAngp1 0 21600000
gd swAng1 +- enAng1 0 stAng1
hxy adj 0 87500 - 0 0 g0w vc
rect g12w g15h g0w g16h
path: M r b A w hd2 cd4 cd2 A g18w dy1 stAng1 swAng1 Z`,
  cloud: `
gd il */ w 2977 21600
gd it */ h 3262 21600
gd ir */ w 17087 21600
gd ib */ h 17337 21600
rect il it ir ib
path w=43200 h=43200: ${CLOUD}
path w=43200 h=43200 fill=none: ${CLOUD_ARCS}`,
  arc: `
av adj1 16200000
av adj2 0
gd stAng pin 0 adj1 21599999
gd enAng pin 0 adj2 21599999
gd sw11 +- enAng 0 stAng
gd sw12 +- sw11 21600000 0
gd swAng ?: sw11 sw11 sw12
gd wt1 sin wd2 stAng
gd ht1 cos hd2 stAng
gd dx1 cat2 wd2 ht1 wt1
gd dy1 sat2 hd2 ht1 wt1
gd wt2 sin wd2 enAng
gd ht2 cos hd2 enAng
gd dx2 cat2 wd2 ht2 wt2
gd dy2 sat2 hd2 ht2 wt2
gd x1 +- hc dx1 0
gd y1 +- vc dy1 0
gd x2 +- hc dx2 0
gd y2 +- vc dy2 0
hpolar - 0 0 adj1 0 21599999 x1 y1
hpolar - 0 0 adj2 0 21599999 x2 y2
rect l t r b
path stroke=false: M x1 y1 A wd2 hd2 stAng swAng L hc vc Z
path fill=none: M x1 y1 A wd2 hd2 stAng swAng`,
  bracketPair: `
av adj 16667
gd a pin 0 adj 50000
gd x1 */ ss a 100000
gd x2 +- r 0 x1
gd y2 +- b 0 x1
gd il */ x1 29289 100000
gd ir +- r 0 il
gd ib +- b 0 il
hxy adj 0 50000 - 0 0 l x1
rect il il ir ib
path stroke=false: M l x1 A x1 x1 cd2 cd4 L x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4 L x1 b A x1 x1 cd4 cd4 Z
path fill=none: M x1 b A x1 x1 cd4 cd4 L l x1 A x1 x1 cd2 cd4 M x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4`,
  bracePair: `
av adj 8333
gd a pin 0 adj 25000
gd x1 */ ss a 100000
gd x2 */ ss a 50000
gd x3 +- r 0 x2
gd x4 +- r 0 x1
gd y2 +- vc 0 x1
gd y3 +- vc x1 0
gd y4 +- b 0 x1
gd it */ x1 29289 100000
gd il +- x1 it 0
gd ir +- r 0 il
gd ib +- b 0 it
hxy - 0 0 adj 0 25000 l x1
rect il il ir ib
path stroke=false: M x2 b A x1 x1 cd4 cd4 L x1 y3 A x1 x1 0 -5400000 A x1 x1 cd4 -5400000 L x1 x1 A x1 x1 cd2 cd4 L x3 t A x1 x1 3cd4 cd4 L x4 y2 A x1 x1 cd2 -5400000 A x1 x1 3cd4 -5400000 L x4 y4 A x1 x1 0 cd4 Z
path fill=none: M x2 b A x1 x1 cd4 cd4 L x1 y3 A x1 x1 0 -5400000 A x1 x1 cd4 -5400000 L x1 x1 A x1 x1 cd2 cd4 M x3 t A x1 x1 3cd4 cd4 L x4 y2 A x1 x1 cd2 -5400000 A x1 x1 3cd4 -5400000 L x4 y4 A x1 x1 0 cd4`,
  leftBracket: `
av adj 8333
gd maxAdj */ 50000 h ss
gd a pin 0 adj maxAdj
gd y1 */ ss a 100000
gd y2 +- b 0 y1
gd dx1 cos w 2700000
gd dy1 sin y1 2700000
gd il +- r 0 dx1
gd it +- y1 0 dy1
gd ib +- b dy1 y1
hxy - 0 0 adj 0 maxAdj l y1
rect il it r ib
path stroke=false: M r b A w y1 cd4 cd4 L l y1 A w y1 cd2 cd4 Z
path fill=none: M r b A w y1 cd4 cd4 L l y1 A w y1 cd2 cd4`,
  rightBracket: `
av adj 8333
gd maxAdj */ 50000 h ss
gd a pin 0 adj maxAdj
gd y1 */ ss a 100000
gd y2 +- b 0 y1
gd dx1 cos w 2700000
gd dy1 sin y1 2700000
gd ir +- l dx1 0
gd it +- y1 0 dy1
gd ib +- b dy1 y1
hxy - 0 0 adj 0 maxAdj r y2
rect l it ir ib
path stroke=false: M l t A w y1 3cd4 cd4 L r y2 A w y1 0 cd4 Z
path fill=none: M l t A w y1 3cd4 cd4 L r y2 A w y1 0 cd4`,
  leftBrace: `
av adj1 8333
av adj2 50000
gd a2 pin 0 adj2 100000
gd q1 +- 100000 0 a2
gd q2 min q1 a2
gd q3 */ q2 1 2
gd maxAdj1 */ q3 h ss
gd a1 pin 0 adj1 maxAdj1
gd y1 */ ss a1 100000
gd y3 */ h a2 100000
gd y4 +- y3 y1 0
gd dx1 cos wd2 2700000
gd dy1 sin y1 2700000
gd il +- r 0 dx1
gd it +- y1 0 dy1
gd ib +- b dy1 y1
hxy - 0 0 adj1 0 maxAdj1 hc y1
hxy - 0 0 adj2 0 100000 l y3
rect il it r ib
path stroke=false: M r b A wd2 y1 cd4 cd4 L hc y4 A wd2 y1 0 -5400000 A wd2 y1 cd4 -5400000 L hc y1 A wd2 y1 cd2 cd4 Z
path fill=none: M r b A wd2 y1 cd4 cd4 L hc y4 A wd2 y1 0 -5400000 A wd2 y1 cd4 -5400000 L hc y1 A wd2 y1 cd2 cd4`,
  rightBrace: `
av adj1 8333
av adj2 50000
gd a2 pin 0 adj2 100000
gd q1 +- 100000 0 a2
gd q2 min q1 a2
gd q3 */ q2 1 2
gd maxAdj1 */ q3 h ss
gd a1 pin 0 adj1 maxAdj1
gd y1 */ ss a1 100000
gd y3 */ h a2 100000
gd y2 +- y3 0 y1
gd y4 +- b 0 y1
gd dx1 cos wd2 2700000
gd dy1 sin y1 2700000
gd ir +- l dx1 0
gd it +- y1 0 dy1
gd ib +- b dy1 y1
hxy - 0 0 adj1 0 maxAdj1 hc y1
hxy - 0 0 adj2 0 100000 r y3
rect l it ir ib
path stroke=false: M l t A wd2 y1 3cd4 cd4 L hc y2 A wd2 y1 cd2 -5400000 A wd2 y1 3cd4 -5400000 L hc y4 A wd2 y1 0 cd4 Z
path fill=none: M l t A wd2 y1 3cd4 cd4 L hc y2 A wd2 y1 cd2 -5400000 A wd2 y1 3cd4 -5400000 L hc y4 A wd2 y1 0 cd4`,

  // ----- block arrows -----------------------------------------------------------------------
  rightArrow: `
av adj1 50000
av adj2 50000
gd maxAdj2 */ 100000 w ss
gd a1 pin 0 adj1 100000
gd a2 pin 0 adj2 maxAdj2
gd dx1 */ ss a2 100000
gd x1 +- r 0 dx1
gd dy1 */ h a1 200000
gd y1 +- vc 0 dy1
gd y2 +- vc dy1 0
gd dx2 */ y1 dx1 hd2
gd x2 +- x1 dx2 0
hxy - 0 0 adj1 0 100000 r y1
hxy adj2 0 maxAdj2 - 0 0 x1 t
rect l y1 x2 y2
path: M l y1 L x1 y1 L x1 t L r vc L x1 b L x1 y2 L l y2 Z`,
  leftArrow: `
av adj1 50000
av adj2 50000
gd maxAdj2 */ 100000 w ss
gd a1 pin 0 adj1 100000
gd a2 pin 0 adj2 maxAdj2
gd dx2 */ ss a2 100000
gd x2 +- l dx2 0
gd dy1 */ h a1 200000
gd y1 +- vc 0 dy1
gd y2 +- vc dy1 0
gd dx1 */ y1 dx2 hd2
gd x1 +- x2 0 dx1
hxy - 0 0 adj1 0 100000 r y1
hxy adj2 0 maxAdj2 - 0 0 x2 t
rect x1 y1 r y2
path: M l vc L x2 t L x2 y1 L r y1 L r y2 L x2 y2 L x2 b Z`,
  upArrow: `
av adj1 50000
av adj2 50000
gd maxAdj2 */ 100000 h ss
gd a1 pin 0 adj1 100000
gd a2 pin 0 adj2 maxAdj2
gd dy2 */ ss a2 100000
gd y2 +- t dy2 0
gd dx1 */ w a1 200000
gd x1 +- hc 0 dx1
gd x2 +- hc dx1 0
gd dy1 */ x1 dy2 wd2
gd y1 +- y2 0 dy1
hxy adj1 0 100000 - 0 0 x1 b
hxy - 0 0 adj2 0 maxAdj2 l y2
rect x1 y1 x2 b
path: M l y2 L hc t L r y2 L x2 y2 L x2 b L x1 b L x1 y2 Z`,
  downArrow: `
av adj1 50000
av adj2 50000
gd maxAdj2 */ 100000 h ss
gd a1 pin 0 adj1 100000
gd a2 pin 0 adj2 maxAdj2
gd dy1 */ ss a2 100000
gd y1 +- b 0 dy1
gd dx1 */ w a1 200000
gd x1 +- hc 0 dx1
gd x2 +- hc dx1 0
gd dy2 */ x1 dy1 wd2
gd y2 +- y1 dy2 0
hxy adj1 0 100000 - 0 0 x1 t
hxy - 0 0 adj2 0 maxAdj2 l y1
rect x1 t x2 y2
path: M l y1 L x1 y1 L x1 t L x2 t L x2 y1 L r y1 L hc b Z`,
  leftRightArrow: `
av adj1 50000
av adj2 50000
gd maxAdj2 */ 50000 w ss
gd a1 pin 0 adj1 100000
gd a2 pin 0 adj2 maxAdj2
gd x2 */ ss a2 100000
gd x3 +- r 0 x2
gd dy */ h a1 200000
gd y1 +- vc 0 dy
gd y2 +- vc dy 0
gd dx1 */ y1 x2 hd2
gd x1 +- x2 0 dx1
gd x4 +- x3 dx1 0
hxy - 0 0 adj1 0 100000 x3 y1
hxy adj2 0 maxAdj2 - 0 0 x2 t
rect x1 y1 x4 y2
path: M l vc L x2 t L x2 y1 L x3 y1 L x3 t L r vc L x3 b L x3 y2 L x2 y2 L x2 b Z`,
  upDownArrow: `
av adj1 50000
av adj2 50000
gd maxAdj2 */ 50000 h ss
gd a1 pin 0 adj1 100000
gd a2 pin 0 adj2 maxAdj2
gd y2 */ ss a2 100000
gd y3 +- b 0 y2
gd dx1 */ w a1 200000
gd x1 +- hc 0 dx1
gd x2 +- hc dx1 0
gd dy1 */ x1 y2 wd2
gd y1 +- y2 0 dy1
gd y4 +- y3 dy1 0
hxy adj1 0 100000 - 0 0 x1 y3
hxy - 0 0 adj2 0 maxAdj2 l y2
rect x1 y1 x2 y4
path: M l y2 L hc t L r y2 L x2 y2 L x2 y3 L r y3 L hc b L l y3 L x1 y3 L x1 y2 Z`,
  quadArrow: `
av adj1 22500
av adj2 22500
av adj3 22500
gd a2 pin 0 adj2 50000
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd q1 +- 100000 0 maxAdj1
gd maxAdj3 */ q1 1 2
gd a3 pin 0 adj3 maxAdj3
gd x1 */ ss a3 100000
gd dx2 */ ss a2 100000
gd x2 +- hc 0 dx2
gd x5 +- hc dx2 0
gd dx3 */ ss a1 200000
gd x3 +- hc 0 dx3
gd x4 +- hc dx3 0
gd x6 +- r 0 x1
gd y2 +- vc 0 dx2
gd y5 +- vc dx2 0
gd y3 +- vc 0 dx3
gd y4 +- vc dx3 0
gd y6 +- b 0 x1
gd il */ dx3 x1 dx2
gd ir +- r 0 il
hxy adj1 0 maxAdj1 - 0 0 x3 x1
hxy adj2 0 50000 - 0 0 x2 t
hxy adj3 0 maxAdj3 - 0 0 r x1
rect il y3 ir y4
path: M l vc L x1 y2 L x1 y3 L x3 y3 L x3 x1 L x2 x1 L hc t L x5 x1 L x4 x1 L x4 y3 L x6 y3 L x6 y2 L r vc L x6 y5 L x6 y4 L x4 y4 L x4 y6 L x5 y6 L hc b L x2 y6 L x3 y6 L x3 y4 L x1 y4 L x1 y5 Z`,
  leftRightUpArrow: `
av adj1 25000
av adj2 25000
av adj3 25000
gd a2 pin 0 adj2 50000
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd q1 +- 100000 0 maxAdj1
gd maxAdj3 */ q1 1 2
gd a3 pin 0 adj3 maxAdj3
gd x1 */ ss a3 100000
gd dx2 */ ss a2 100000
gd x2 +- hc 0 dx2
gd x5 +- hc dx2 0
gd dx3 */ ss a1 200000
gd x3 +- hc 0 dx3
gd x4 +- hc dx3 0
gd x6 +- r 0 x1
gd dy2 */ ss a2 50000
gd y2 +- b 0 dy2
gd y4 +- b 0 dx2
gd y3 +- y4 0 dx3
gd y5 +- y4 dx3 0
gd il */ dx3 x1 dx2
gd ir +- r 0 il
hxy adj1 0 maxAdj1 - 0 0 x3 x1
hxy adj2 0 50000 - 0 0 x2 t
hxy adj3 0 maxAdj3 - 0 0 r x1
rect il y3 ir y5
path: M l y4 L x1 y2 L x1 y3 L x3 y3 L x3 x1 L x2 x1 L hc t L x5 x1 L x4 x1 L x4 y3 L x6 y3 L x6 y2 L r y4 L x6 b L x6 y5 L x1 y5 L x1 b Z`,
  bentArrow: `
av adj1 25000
av adj2 25000
av adj3 25000
av adj4 43750
gd a2 pin 0 adj2 50000
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd a3 pin 0 adj3 50000
gd th */ ss a1 100000
gd aw2 */ ss a2 100000
gd th2 */ th 1 2
gd dh2 +- aw2 0 th2
gd ah */ ss a3 100000
gd bw +- r 0 ah
gd bh +- b 0 dh2
gd bs min bw bh
gd maxAdj4 */ 100000 bs ss
gd a4 pin 0 adj4 maxAdj4
gd bd */ ss a4 100000
gd bd3 +- bd 0 th
gd bd2 max bd3 0
gd x3 +- th bd2 0
gd x4 +- r 0 ah
gd y3 +- dh2 th 0
gd y4 +- y3 dh2 0
gd y5 +- dh2 bd 0
gd y6 +- y3 bd2 0
hxy adj1 0 maxAdj1 - 0 0 th b
hxy - 0 0 adj2 0 50000 r y4
hxy adj3 0 50000 - 0 0 x4 t
hxy adj4 0 maxAdj4 - 0 0 bd t
rect l y3 x4 b
path: M l b L l y5 A bd bd cd2 cd4 L x4 dh2 L x4 t L r aw2 L x4 y4 L x4 y3 L x3 y3 A bd2 bd2 3cd4 -5400000 L th b Z`,
  uturnArrow: `
av adj1 25000
av adj2 25000
av adj3 25000
av adj4 43750
av adj5 75000
gd a2 pin 0 adj2 25000
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd q2 */ a1 ss h
gd q3 +- 100000 0 q2
gd maxAdj3 */ q3 h ss
gd a3 pin 0 adj3 maxAdj3
gd q1 +- a3 a1 0
gd minAdj5 */ q1 ss h
gd a5 pin minAdj5 adj5 100000
gd th */ ss a1 100000
gd aw2 */ ss a2 100000
gd th2 */ th 1 2
gd dh2 +- aw2 0 th2
gd y5 */ h a5 100000
gd ah */ ss a3 100000
gd y4 +- y5 0 ah
gd x9 +- r 0 dh2
gd bw */ x9 1 2
gd bs min bw y4
gd maxAdj4 */ bs 100000 ss
gd a4 pin 0 adj4 maxAdj4
gd bd */ ss a4 100000
gd bd3 +- bd 0 th
gd bd2 max bd3 0
gd x3 +- th bd2 0
gd x8 +- r 0 aw2
gd x6 +- x8 0 aw2
gd x7 +- x6 dh2 0
gd x4 +- x9 0 bd
gd x5 +- x7 0 bd2
hxy adj1 0 maxAdj1 - 0 0 th b
hxy adj2 0 25000 - 0 0 x6 b
hxy - 0 0 adj3 0 maxAdj3 x6 y4
hxy adj4 0 maxAdj4 - 0 0 bd t
hxy - 0 0 adj5 minAdj5 100000 r y5
rect l t x9 y4
path: M l b L l bd A bd bd cd2 cd4 L x4 t A bd bd 3cd4 cd4 L x9 y4 L r y4 L x8 y5 L x6 y4 L x7 y4 L x7 x3 A bd2 bd2 0 -5400000 L x3 th A bd2 bd2 3cd4 -5400000 L th b Z`,
  leftUpArrow: `
av adj1 25000
av adj2 25000
av adj3 25000
gd a2 pin 0 adj2 50000
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd maxAdj3 +- 100000 0 maxAdj1
gd a3 pin 0 adj3 maxAdj3
gd x1 */ ss a3 100000
gd dx2 */ ss a2 50000
gd x2 +- r 0 dx2
gd y2 +- b 0 dx2
gd dx4 */ ss a2 100000
gd x4 +- r 0 dx4
gd y4 +- b 0 dx4
gd dx3 */ ss a1 200000
gd x3 +- x4 0 dx3
gd x5 +- x4 dx3 0
gd y3 +- y4 0 dx3
gd y5 +- y4 dx3 0
gd il */ dx3 x1 dx4
hxy - 0 0 adj1 0 maxAdj1 x3 y3
hxy adj2 0 50000 - 0 0 x2 t
hxy adj3 0 maxAdj3 - 0 0 x3 x1
rect il y3 x4 y5
path: M l y4 L x1 y2 L x1 y3 L x3 y3 L x3 x1 L x2 x1 L x4 t L r x1 L x5 x1 L x5 y5 L x1 y5 L x1 b Z`,
  bentUpArrow: `
av adj1 25000
av adj2 25000
av adj3 25000
gd a1 pin 0 adj1 50000
gd a2 pin 0 adj2 50000
gd a3 pin 0 adj3 50000
gd y1 */ ss a3 100000
gd dx1 */ ss a2 50000
gd x1 +- r 0 dx1
gd dx3 */ ss a2 100000
gd x3 +- r 0 dx3
gd dx2 */ ss a1 200000
gd x2 +- x3 0 dx2
gd x4 +- x3 dx2 0
gd dy2 */ ss a1 100000
gd y2 +- b 0 dy2
hxy - 0 0 adj1 0 50000 l y2
hxy adj2 0 50000 - 0 0 x1 t
hxy - 0 0 adj3 0 50000 x3 y1
rect l y2 x4 b
path: M l y2 L x2 y2 L x2 y1 L x1 y1 L x3 t L r y1 L x4 y1 L x4 b L l b Z`,
  curvedRightArrow: `
av adj1 25000
av adj2 50000
av adj3 25000
gd maxAdj2 */ 50000 h ss
gd a2 pin 0 adj2 maxAdj2
gd a1 pin 0 adj1 a2
gd th */ ss a1 100000
gd aw */ ss a2 100000
gd q1 +/ th aw 4
gd hR +- hd2 0 q1
gd q7 */ hR 2 1
gd q8 */ q7 q7 1
gd q9 */ th th 1
gd q10 +- q8 0 q9
gd q11 sqrt q10
gd idx */ q11 w q7
gd maxAdj3 */ 100000 idx ss
gd a3 pin 0 adj3 maxAdj3
gd ah */ ss a3 100000
gd y3 +- hR th 0
gd q2 */ w w 1
gd q3 */ ah ah 1
gd q4 +- q2 0 q3
gd q5 sqrt q4
gd dy */ q5 hR w
gd y5 +- hR dy 0
gd y7 +- y3 dy 0
gd q6 +- aw 0 th
gd dh */ q6 1 2
gd y4 +- y5 0 dh
gd y8 +- y7 dh 0
gd aw2 */ aw 1 2
gd y6 +- b 0 aw2
gd x1 +- r 0 ah
gd swAng at2 ah dy
gd stAng +- cd2 0 swAng
gd mswAng +- 0 0 swAng
gd ix +- r 0 idx
gd iy +/ hR y3 2
gd q12 */ th 1 2
gd dang2 at2 idx q12
gd swAng2 +- dang2 0 cd4
gd swAng3 +- cd4 dang2 0
gd stAng3 +- cd2 0 dang2
hxy - 0 0 adj1 0 a2 x1 y5
hxy - 0 0 adj2 0 maxAdj2 r y4
hxy adj3 0 maxAdj3 - 0 0 x1 b
rect l t r b
path stroke=false: M l hR A w hR cd2 mswAng L x1 y4 L r y6 L x1 y8 L x1 y7 A w hR stAng swAng Z
path stroke=false fill=darkenLess: M r th A w hR 3cd4 swAng2 A w hR stAng3 swAng3 Z
path fill=none: M l hR A w hR cd2 mswAng L x1 y4 L r y6 L x1 y8 L x1 y7 A w hR stAng swAng L l hR A w hR cd2 cd4 L r th A w hR 3cd4 swAng2`,
  curvedLeftArrow: `
av adj1 25000
av adj2 50000
av adj3 25000
gd maxAdj2 */ 50000 h ss
gd a2 pin 0 adj2 maxAdj2
gd a1 pin 0 adj1 a2
gd th */ ss a1 100000
gd aw */ ss a2 100000
gd q1 +/ th aw 4
gd hR +- hd2 0 q1
gd q7 */ hR 2 1
gd q8 */ q7 q7 1
gd q9 */ th th 1
gd q10 +- q8 0 q9
gd q11 sqrt q10
gd iDx */ q11 w q7
gd maxAdj3 */ 100000 iDx ss
gd a3 pin 0 adj3 maxAdj3
gd ah */ ss a3 100000
gd y3 +- hR th 0
gd q2 */ w w 1
gd q3 */ ah ah 1
gd q4 +- q2 0 q3
gd q5 sqrt q4
gd dy */ q5 hR w
gd y5 +- hR dy 0
gd y7 +- y3 dy 0
gd q6 +- aw 0 th
gd dh */ q6 1 2
gd y4 +- y5 0 dh
gd y8 +- y7 dh 0
gd aw2 */ aw 1 2
gd y6 +- b 0 aw2
gd x1 +- l ah 0
gd swAng at2 ah dy
gd q12 */ th 1 2
gd dang2 at2 iDx q12
gd swAng2 +- dang2 0 swAng
gd swAng3 +- swAng dang2 0
gd stAng3 +- 0 0 dang2
hxy - 0 0 adj1 0 a2 x1 y5
hxy - 0 0 adj2 0 maxAdj2 r y4
hxy adj3 0 maxAdj3 - 0 0 x1 b
rect l t r b
path stroke=false: M l y6 L x1 y4 L x1 y5 A w hR swAng swAng2 A w hR stAng3 swAng3 L x1 y8 Z
path stroke=false fill=darkenLess: M r y3 A w hR 0 -5400000 L l t A w hR 3cd4 cd4 Z
path fill=none: M r y3 A w hR 0 -5400000 L l t A w hR 3cd4 cd4 L r y3 A w hR 0 swAng L x1 y7 L x1 y8 L l y6 L x1 y4 L x1 y5 A w hR swAng swAng2`,
  curvedUpArrow: `
av adj1 25000
av adj2 50000
av adj3 25000
gd maxAdj2 */ 50000 w ss
gd a2 pin 0 adj2 maxAdj2
gd a1 pin 0 adj1 100000
gd th */ ss a1 100000
gd aw */ ss a2 100000
gd q1 +/ th aw 4
gd wR +- wd2 0 q1
gd q7 */ wR 2 1
gd q8 */ q7 q7 1
gd q9 */ th th 1
gd q10 +- q8 0 q9
gd q11 sqrt q10
gd idy */ q11 h q7
gd maxAdj3 */ 100000 idy ss
gd a3 pin 0 adj3 maxAdj3
gd ah */ ss adj3 100000
gd x3 +- wR th 0
gd q2 */ h h 1
gd q3 */ ah ah 1
gd q4 +- q2 0 q3
gd q5 sqrt q4
gd dx */ q5 wR h
gd x5 +- wR dx 0
gd x7 +- x3 dx 0
gd q6 +- aw 0 th
gd dh */ q6 1 2
gd x4 +- x5 0 dh
gd x8 +- x7 dh 0
gd aw2 */ aw 1 2
gd x6 +- r 0 aw2
gd y1 +- t ah 0
gd swAng at2 ah dx
gd mswAng +- 0 0 swAng
gd iy +- t idy 0
gd ix +/ wR x3 2
gd q12 */ th 1 2
gd dang2 at2 idy q12
gd swAng2 +- dang2 0 swAng
gd mswAng2 +- 0 0 swAng2
gd stAng3 +- cd4 0 swAng
gd swAng3 +- swAng dang2 0
gd stAng2 +- cd4 0 dang2
hxy adj1 0 100000 - 0 0 x7 y1
hxy adj2 0 maxAdj2 - 0 0 x4 b
hxy - 0 0 adj3 0 maxAdj3 r y1
rect l t r b
path stroke=false: M x6 t L x8 y1 L x7 y1 A wR h stAng3 swAng3 A wR h stAng2 swAng2 L x5 y1 L x4 y1 Z
path stroke=false fill=darkenLess: M wR b A wR h cd4 cd4 L th t A wR h cd2 -5400000 Z
path fill=none: M ix iy A wR h stAng2 swAng2 L x5 y1 L x4 y1 L x6 t L x8 y1 L x7 y1 A wR h stAng3 swAng L wR b A wR h cd4 cd4 L th t A wR h cd2 -5400000`,
  curvedDownArrow: `
av adj1 25000
av adj2 50000
av adj3 25000
gd maxAdj2 */ 50000 w ss
gd a2 pin 0 adj2 maxAdj2
gd a1 pin 0 adj1 100000
gd th */ ss a1 100000
gd aw */ ss a2 100000
gd q1 +/ th aw 4
gd wR +- wd2 0 q1
gd q7 */ wR 2 1
gd q8 */ q7 q7 1
gd q9 */ th th 1
gd q10 +- q8 0 q9
gd q11 sqrt q10
gd idy */ q11 h q7
gd maxAdj3 */ 100000 idy ss
gd a3 pin 0 adj3 maxAdj3
gd ah */ ss adj3 100000
gd x3 +- wR th 0
gd q2 */ h h 1
gd q3 */ ah ah 1
gd q4 +- q2 0 q3
gd q5 sqrt q4
gd dx */ q5 wR h
gd x5 +- wR dx 0
gd x7 +- x3 dx 0
gd q6 +- aw 0 th
gd dh */ q6 1 2
gd x4 +- x5 0 dh
gd x8 +- x7 dh 0
gd aw2 */ aw 1 2
gd x6 +- r 0 aw2
gd y1 +- b 0 ah
gd swAng at2 ah dx
gd mswAng +- 0 0 swAng
gd iy +- b 0 idy
gd ix +/ wR x3 2
gd q12 */ th 1 2
gd dang2 at2 idy q12
gd stAng +- 3cd4 swAng 0
gd stAng2 +- 3cd4 0 dang2
gd swAng2 +- dang2 0 cd4
gd swAng3 +- cd4 dang2 0
hxy adj1 0 100000 - 0 0 x7 y1
hxy adj2 0 maxAdj2 - 0 0 x4 t
hxy - 0 0 adj3 0 maxAdj3 r y1
rect l t r b
path stroke=false: M x6 b L x4 y1 L x5 y1 A wR h stAng mswAng L x3 t A wR h 3cd4 swAng L x8 y1 Z
path stroke=false fill=darkenLess: M ix iy A wR h stAng2 swAng2 L l b A wR h cd2 swAng3 Z
path fill=none: M x6 b L x4 y1 L x5 y1 A wR h stAng mswAng L x3 t A wR h 3cd4 swAng L x8 y1 L x6 b M ix iy A wR h stAng2 swAng2 L l b A wR h cd2 cd4`,
  stripedRightArrow: `
av adj1 50000
av adj2 50000
gd maxAdj2 */ 84375 w ss
gd a1 pin 0 adj1 100000
gd a2 pin 0 adj2 maxAdj2
gd x4 */ ss 5 32
gd dx5 */ ss a2 100000
gd x5 +- r 0 dx5
gd dy1 */ h a1 200000
gd y1 +- vc 0 dy1
gd y2 +- vc dy1 0
gd dx6 */ dy1 dx5 hd2
gd x6 +- r 0 dx6
hxy - 0 0 adj1 0 100000 l y1
hxy adj2 0 maxAdj2 - 0 0 x5 t
rect x4 y1 x6 y2
path: M l y1 L ssd32 y1 L ssd32 y2 L l y2 Z M ssd16 y1 L ssd8 y1 L ssd8 y2 L ssd16 y2 Z M x4 y1 L x5 y1 L x5 t L r vc L x5 b L x5 y2 L x4 y2 Z`,
  notchedRightArrow: `
av adj1 50000
av adj2 50000
gd maxAdj2 */ 100000 w ss
gd a1 pin 0 adj1 100000
gd a2 pin 0 adj2 maxAdj2
gd dx2 */ ss a2 100000
gd x2 +- r 0 dx2
gd dy1 */ h a1 200000
gd y1 +- vc 0 dy1
gd y2 +- vc dy1 0
gd x1 */ dy1 dx2 hd2
gd x3 +- r 0 x1
hxy - 0 0 adj1 0 100000 r y1
hxy adj2 0 maxAdj2 - 0 0 x2 t
rect x1 y1 x3 y2
path: M l y1 L x2 y1 L x2 t L r vc L x2 b L x2 y2 L l y2 L x1 vc Z`,
  homePlate: `
av adj 50000
gd maxAdj */ 100000 w ss
gd a pin 0 adj maxAdj
gd dx1 */ ss a 100000
gd x1 +- r 0 dx1
gd ir +/ x1 r 2
hxy adj 0 maxAdj - 0 0 x1 t
rect l t ir b
path: M l t L x1 t L r vc L x1 b L l b Z`,
  chevron: `
av adj 50000
gd maxAdj */ 100000 w ss
gd a pin 0 adj maxAdj
gd x1 */ ss a 100000
gd x2 +- r 0 x1
gd dx +- x2 0 x1
gd il ?: dx x1 l
gd ir ?: dx x2 r
hxy adj 0 maxAdj - 0 0 x2 t
rect il t ir b
path: M l t L x2 t L r vc L x2 b L l b L x1 vc Z`,
  rightArrowCallout: `
av adj1 25000
av adj2 25000
av adj3 25000
av adj4 64977
gd maxAdj2 */ 50000 h ss
gd a2 pin 0 adj2 maxAdj2
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd maxAdj3 */ 100000 w ss
gd a3 pin 0 adj3 maxAdj3
gd q2 */ a3 ss w
gd maxAdj4 +- 100000 0 q2
gd a4 pin 0 adj4 maxAdj4
gd dy1 */ ss a2 100000
gd dy2 */ ss a1 200000
gd y1 +- vc 0 dy1
gd y2 +- vc 0 dy2
gd y3 +- vc dy2 0
gd y4 +- vc dy1 0
gd dx3 */ ss a3 100000
gd x3 +- r 0 dx3
gd x2 */ w a4 100000
hxy - 0 0 adj1 0 maxAdj1 x3 y2
hxy - 0 0 adj2 0 maxAdj2 r y1
hxy adj3 0 maxAdj3 - 0 0 x3 t
hxy adj4 0 maxAdj4 - 0 0 x2 b
rect l t x2 b
path: M l t L x2 t L x2 y2 L x3 y2 L x3 y1 L r vc L x3 y4 L x3 y3 L x2 y3 L x2 b L l b Z`,
  leftArrowCallout: `
av adj1 25000
av adj2 25000
av adj3 25000
av adj4 64977
gd maxAdj2 */ 50000 h ss
gd a2 pin 0 adj2 maxAdj2
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd maxAdj3 */ 100000 w ss
gd a3 pin 0 adj3 maxAdj3
gd q2 */ a3 ss w
gd maxAdj4 +- 100000 0 q2
gd a4 pin 0 adj4 maxAdj4
gd dy1 */ ss a2 100000
gd dy2 */ ss a1 200000
gd y1 +- vc 0 dy1
gd y2 +- vc 0 dy2
gd y3 +- vc dy2 0
gd y4 +- vc dy1 0
gd x1 */ ss a3 100000
gd dx2 */ w a4 100000
gd x2 +- r 0 dx2
hxy - 0 0 adj1 0 maxAdj1 x1 y2
hxy - 0 0 adj2 0 maxAdj2 l y1
hxy adj3 0 maxAdj3 - 0 0 x1 t
hxy adj4 0 maxAdj4 - 0 0 x2 b
rect x2 t r b
path: M l vc L x1 y1 L x1 y2 L x2 y2 L x2 t L r t L r b L x2 b L x2 y3 L x1 y3 L x1 y4 Z`,
  upArrowCallout: `
av adj1 25000
av adj2 25000
av adj3 25000
av adj4 64977
gd maxAdj2 */ 50000 w ss
gd a2 pin 0 adj2 maxAdj2
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd maxAdj3 */ 100000 h ss
gd a3 pin 0 adj3 maxAdj3
gd q2 */ a3 ss h
gd maxAdj4 +- 100000 0 q2
gd a4 pin 0 adj4 maxAdj4
gd dx1 */ ss a2 100000
gd dx2 */ ss a1 200000
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc dx2 0
gd x4 +- hc dx1 0
gd y1 */ ss a3 100000
gd dy2 */ h a4 100000
gd y2 +- b 0 dy2
hxy adj1 0 maxAdj1 - 0 0 x2 y1
hxy adj2 0 maxAdj2 - 0 0 x1 t
hxy - 0 0 adj3 0 maxAdj3 r y1
hxy - 0 0 adj4 0 maxAdj4 l y2
rect l y2 r b
path: M l y2 L x2 y2 L x2 y1 L x1 y1 L hc t L x4 y1 L x3 y1 L x3 y2 L r y2 L r b L l b Z`,
  downArrowCallout: `
av adj1 25000
av adj2 25000
av adj3 25000
av adj4 64977
gd maxAdj2 */ 50000 w ss
gd a2 pin 0 adj2 maxAdj2
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd maxAdj3 */ 100000 h ss
gd a3 pin 0 adj3 maxAdj3
gd q2 */ a3 ss h
gd maxAdj4 +- 100000 0 q2
gd a4 pin 0 adj4 maxAdj4
gd dx1 */ ss a2 100000
gd dx2 */ ss a1 200000
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc dx2 0
gd x4 +- hc dx1 0
gd dy3 */ ss a3 100000
gd y3 +- b 0 dy3
gd y2 */ h a4 100000
hxy adj1 0 maxAdj1 - 0 0 x2 y3
hxy adj2 0 maxAdj2 - 0 0 x1 b
hxy - 0 0 adj3 0 maxAdj3 r y3
hxy - 0 0 adj4 0 maxAdj4 l y2
rect l t r y2
path: M l t L r t L r y2 L x3 y2 L x3 y3 L x4 y3 L hc b L x1 y3 L x2 y3 L x2 y2 L l y2 Z`,
  leftRightArrowCallout: `
av adj1 25000
av adj2 25000
av adj3 25000
av adj4 48123
gd maxAdj2 */ 50000 h ss
gd a2 pin 0 adj2 maxAdj2
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd maxAdj3 */ 50000 w ss
gd a3 pin 0 adj3 maxAdj3
gd q2 */ a3 ss wd2
gd maxAdj4 +- 100000 0 q2
gd a4 pin 0 adj4 maxAdj4
gd dy1 */ ss a2 100000
gd dy2 */ ss a1 200000
gd y1 +- vc 0 dy1
gd y2 +- vc 0 dy2
gd y3 +- vc dy2 0
gd y4 +- vc dy1 0
gd x1 */ ss a3 100000
gd x4 +- r 0 x1
gd dx2 */ w a4 200000
gd x2 +- hc 0 dx2
gd x3 +- hc dx2 0
hxy - 0 0 adj1 0 maxAdj1 x1 y2
hxy - 0 0 adj2 0 maxAdj2 l y1
hxy adj3 0 maxAdj3 - 0 0 x1 t
hxy adj4 0 maxAdj4 - 0 0 x2 b
rect x2 t x3 b
path: M l vc L x1 y1 L x1 y2 L x2 y2 L x2 t L x3 t L x3 y2 L x4 y2 L x4 y1 L r vc L x4 y4 L x4 y3 L x3 y3 L x3 b L x2 b L x2 y3 L x1 y3 L x1 y4 Z`,
  upDownArrowCallout: `
av adj1 25000
av adj2 25000
av adj3 25000
av adj4 48123
gd maxAdj2 */ 50000 w ss
gd a2 pin 0 adj2 maxAdj2
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd maxAdj3 */ 50000 h ss
gd a3 pin 0 adj3 maxAdj3
gd q2 */ a3 ss hd2
gd maxAdj4 +- 100000 0 q2
gd a4 pin 0 adj4 maxAdj4
gd dx1 */ ss a2 100000
gd dx2 */ ss a1 200000
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc dx2 0
gd x4 +- hc dx1 0
gd y1 */ ss a3 100000
gd y4 +- b 0 y1
gd dy2 */ h a4 200000
gd y2 +- vc 0 dy2
gd y3 +- vc dy2 0
hxy adj1 0 maxAdj1 - 0 0 x2 y1
hxy adj2 0 maxAdj2 - 0 0 x1 t
hxy - 0 0 adj3 0 maxAdj3 r y1
hxy - 0 0 adj4 0 maxAdj4 l y2
rect l y2 r y3
path: M l y2 L x2 y2 L x2 y1 L x1 y1 L hc t L x4 y1 L x3 y1 L x3 y2 L r y2 L r y3 L x3 y3 L x3 y4 L x4 y4 L hc b L x1 y4 L x2 y4 L x2 y3 L l y3 Z`,
  circularArrow: `
av adj1 12500
av adj2 1142319
av adj3 20457681
av adj4 10800000
av adj5 12500
gd a5 pin 0 adj5 25000
gd maxAdj1 */ a5 2 1
gd a1 pin 0 adj1 maxAdj1
gd enAng pin 1 adj3 21599999
gd stAng pin 0 adj4 21599999
gd th */ ss a1 100000
gd thh */ ss a5 100000
gd th2 */ th 1 2
gd rw1 +- wd2 th2 thh
gd rh1 +- hd2 th2 thh
gd rw2 +- rw1 0 th
gd rh2 +- rh1 0 th
gd rw3 +- rw2 th2 0
gd rh3 +- rh2 th2 0
gd wtH sin rw3 enAng
gd htH cos rh3 enAng
gd dxH cat2 rw3 htH wtH
gd dyH sat2 rh3 htH wtH
gd xH +- hc dxH 0
gd yH +- vc dyH 0
gd rI min rw2 rh2
gd u1 */ dxH dxH 1
gd u2 */ dyH dyH 1
gd u3 */ rI rI 1
gd u4 +- u1 0 u3
gd u5 +- u2 0 u3
gd u6 */ u4 u5 u1
gd u7 */ u6 1 u2
gd u8 +- 1 0 u7
gd u9 sqrt u8
gd u10 */ u4 1 dxH
gd u11 */ u10 1 dyH
gd u12 +/ 1 u9 u11
gd u13 at2 1 u12
gd u14 +- u13 21600000 0
gd u15 ?: u13 u13 u14
gd u16 +- u15 0 enAng
gd u17 +- u16 21600000 0
gd u18 ?: u16 u16 u17
gd u19 +- u18 0 cd2
gd u20 +- u18 0 21600000
gd u21 ?: u19 u20 u18
gd maxAng abs u21
gd aAng pin 0 adj2 maxAng
gd ptAng +- enAng aAng 0
gd wtA sin rw3 ptAng
gd htA cos rh3 ptAng
gd dxA cat2 rw3 htA wtA
gd dyA sat2 rh3 htA wtA
gd xA +- hc dxA 0
gd yA +- vc dyA 0
gd wtE sin rw1 stAng
gd htE cos rh1 stAng
gd dxE cat2 rw1 htE wtE
gd dyE sat2 rh1 htE wtE
gd xE +- hc dxE 0
gd yE +- vc dyE 0
gd dxG cos thh ptAng
gd dyG sin thh ptAng
gd xG +- xH dxG 0
gd yG +- yH dyG 0
gd dxB cos thh ptAng
gd dyB sin thh ptAng
gd xB +- xH 0 dxB
gd yB +- yH 0 dyB
gd sx1 +- xB 0 hc
gd sy1 +- yB 0 vc
gd sx2 +- xG 0 hc
gd sy2 +- yG 0 vc
gd rO min rw1 rh1
gd x1O */ sx1 rO rw1
gd y1O */ sy1 rO rh1
gd x2O */ sx2 rO rw1
gd y2O */ sy2 rO rh1
gd dxO +- x2O 0 x1O
gd dyO +- y2O 0 y1O
gd dO mod dxO dyO 0
gd q1 */ x1O y2O 1
gd q2 */ x2O y1O 1
gd DO +- q1 0 q2
gd q3 */ rO rO 1
gd q4 */ dO dO 1
gd q5 */ q3 q4 1
gd q6 */ DO DO 1
gd q7 +- q5 0 q6
gd q8 max q7 0
gd sdelO sqrt q8
gd ndyO */ dyO -1 1
gd sdyO ?: ndyO -1 1
gd q9 */ sdyO dxO 1
gd q10 */ q9 sdelO 1
gd q11 */ DO dyO 1
gd dxF1 +/ q11 q10 q4
gd q12 +- q11 0 q10
gd dxF2 */ q12 1 q4
gd adyO abs dyO
gd q13 */ adyO sdelO 1
gd q14 */ DO dxO -1
gd dyF1 +/ q14 q13 q4
gd q15 +- q14 0 q13
gd dyF2 */ q15 1 q4
gd q16 +- x2O 0 dxF1
gd q17 +- x2O 0 dxF2
gd q18 +- y2O 0 dyF1
gd q19 +- y2O 0 dyF2
gd q20 mod q16 q18 0
gd q21 mod q17 q19 0
gd q22 +- q21 0 q20
gd dxF ?: q22 dxF1 dxF2
gd dyF ?: q22 dyF1 dyF2
gd sdxF */ dxF rw1 rO
gd sdyF */ dyF rh1 rO
gd xF +- hc sdxF 0
gd yF +- vc sdyF 0
gd x1I */ sx1 rI rw2
gd y1I */ sy1 rI rh2
gd x2I */ sx2 rI rw2
gd y2I */ sy2 rI rh2
gd dxI +- x2I 0 x1I
gd dyI +- y2I 0 y1I
gd dI mod dxI dyI 0
gd v1 */ x1I y2I 1
gd v2 */ x2I y1I 1
gd DI +- v1 0 v2
gd v3 */ rI rI 1
gd v4 */ dI dI 1
gd v5 */ v3 v4 1
gd v6 */ DI DI 1
gd v7 +- v5 0 v6
gd v8 max v7 0
gd sdelI sqrt v8
gd v9 */ sdyO dxI 1
gd v10 */ v9 sdelI 1
gd v11 */ DI dyI 1
gd dxC1 +/ v11 v10 v4
gd v12 +- v11 0 v10
gd dxC2 */ v12 1 v4
gd adyI abs dyI
gd v13 */ adyI sdelI 1
gd v14 */ DI dxI -1
gd dyC1 +/ v14 v13 v4
gd v15 +- v14 0 v13
gd dyC2 */ v15 1 v4
gd v16 +- x1I 0 dxC1
gd v17 +- x1I 0 dxC2
gd v18 +- y1I 0 dyC1
gd v19 +- y1I 0 dyC2
gd v20 mod v16 v18 0
gd v21 mod v17 v19 0
gd v22 +- v21 0 v20
gd dxC ?: v22 dxC1 dxC2
gd dyC ?: v22 dyC1 dyC2
gd sdxC */ dxC rw2 rI
gd sdyC */ dyC rh2 rI
gd xC +- hc sdxC 0
gd yC +- vc sdyC 0
gd ist0 at2 sdxC sdyC
gd ist1 +- ist0 21600000 0
gd istAng ?: ist0 ist0 ist1
gd isw1 +- stAng 0 istAng
gd isw2 +- isw1 0 21600000
gd iswAng ?: isw1 isw2 isw1
gd p1 +- xF 0 xC
gd p2 +- yF 0 yC
gd p3 mod p1 p2 0
gd p4 */ p3 1 2
gd p5 +- p4 0 thh
gd xGp ?: p5 xF xG
gd yGp ?: p5 yF yG
gd xBp ?: p5 xC xB
gd yBp ?: p5 yC yB
gd en0 at2 sdxF sdyF
gd en1 +- en0 21600000 0
gd en2 ?: en0 en0 en1
gd sw0 +- en2 0 stAng
gd sw1 +- sw0 21600000 0
gd swAng ?: sw0 sw0 sw1
gd wtI sin rw3 stAng
gd htI cos rh3 stAng
gd dxI2 cat2 rw3 htI wtI
gd dyI2 sat2 rh3 htI wtI
gd xI +- hc dxI2 0
gd yI +- vc dyI2 0
gd aI +- stAng 0 cd4
gd aA +- ptAng cd4 0
gd aB +- ptAng cd2 0
gd idx cos rw1 2700000
gd idy sin rh1 2700000
gd il +- hc 0 idx
gd ir +- hc idx 0
gd it +- vc 0 idy
gd ib +- vc idy 0
hpolar - 0 0 adj2 0 maxAng xA yA
hpolar - 0 0 adj4 0 21599999 xE yE
hpolar adj1 0 maxAdj1 adj3 0 21599999 xF yF
hpolar adj5 0 25000 - 0 0 xB yB
rect il it ir ib
path: M xE yE A rw1 rh1 stAng swAng L xGp yGp L xA yA L xBp yBp L xC yC A rw2 rh2 istAng iswAng Z`,
  quadArrowCallout: `
av adj1 18515
av adj2 18515
av adj3 18515
av adj4 48123
gd a2 pin 0 adj2 50000
gd maxAdj1 */ a2 2 1
gd a1 pin 0 adj1 maxAdj1
gd maxAdj3 +- 50000 0 a2
gd a3 pin 0 adj3 maxAdj3
gd q2 */ a3 2 1
gd maxAdj4 +- 100000 0 q2
gd a4 pin a1 adj4 maxAdj4
gd dx2 */ ss a2 100000
gd dx3 */ ss a1 200000
gd ah */ ss a3 100000
gd dx1 */ w a4 200000
gd dy1 */ h a4 200000
gd x8 +- r 0 ah
gd x2 +- hc 0 dx1
gd x7 +- hc dx1 0
gd x3 +- hc 0 dx2
gd x6 +- hc dx2 0
gd x4 +- hc 0 dx3
gd x5 +- hc dx3 0
gd y8 +- b 0 ah
gd y2 +- vc 0 dy1
gd y7 +- vc dy1 0
gd y3 +- vc 0 dx2
gd y6 +- vc dx2 0
gd y4 +- vc 0 dx3
gd y5 +- vc dx3 0
hxy adj1 0 maxAdj1 - 0 0 x4 ah
hxy adj2 0 50000 - 0 0 x3 t
hxy - 0 0 adj3 0 maxAdj3 r ah
hxy - 0 0 adj4 a1 maxAdj4 l y2
rect x2 y2 x7 y7
path: M l vc L ah y3 L ah y4 L x2 y4 L x2 y2 L x4 y2 L x4 ah L x3 ah L hc t L x6 ah L x5 ah L x5 y2 L x7 y2 L x7 y4 L x8 y4 L x8 y3 L r vc L x8 y6 L x8 y5 L x7 y5 L x7 y7 L x5 y7 L x5 y8 L x6 y8 L hc b L x3 y8 L x4 y8 L x4 y7 L x2 y7 L x2 y5 L ah y5 L ah y6 Z`,

  // ----- equation shapes --------------------------------------------------------------------
  mathPlus: `
av adj1 23520
gd a1 pin 0 adj1 73490
gd dx1 */ w 73490 200000
gd dy1 */ h 73490 200000
gd dx2 */ ss a1 200000
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc dx2 0
gd x4 +- hc dx1 0
gd y1 +- vc 0 dy1
gd y2 +- vc 0 dx2
gd y3 +- vc dx2 0
gd y4 +- vc dy1 0
hxy adj1 0 73490 - 0 0 l y2
rect x1 y2 x4 y3
path: M x1 y2 L x2 y2 L x2 y1 L x3 y1 L x3 y2 L x4 y2 L x4 y3 L x3 y3 L x3 y4 L x2 y4 L x2 y3 L x1 y3 Z`,
  mathMinus: `
av adj1 23520
gd a1 pin 0 adj1 100000
gd dy1 */ h a1 200000
gd dx1 */ w 73490 200000
gd y1 +- vc 0 dy1
gd y2 +- vc dy1 0
gd x1 +- hc 0 dx1
gd x2 +- hc dx1 0
hxy - 0 0 adj1 0 100000 hc y1
rect x1 y1 x2 y2
path: M x1 y1 L x2 y1 L x2 y2 L x1 y2 Z`,
  mathMultiply: `
av adj1 23520
gd a1 pin 0 adj1 51965
gd th */ ss a1 100000
gd a at2 w h
gd sa sin 1 a
gd ca cos 1 a
gd ta tan 1 a
gd dl mod w h 0
gd rw */ dl 51965 100000
gd lM +- dl 0 rw
gd xM */ ca lM 2
gd yM */ sa lM 2
gd dxAM */ sa th 2
gd dyAM */ ca th 2
gd xA +- xM 0 dxAM
gd yA +- yM dyAM 0
gd xB +- xM dxAM 0
gd yB +- yM 0 dyAM
gd xBC +- hc 0 xB
gd yBC */ xBC ta 1
gd yC +- yBC yB 0
gd xD +- r 0 xB
gd xE +- r 0 xA
gd yFE +- vc 0 yA
gd xFE */ yFE 1 ta
gd xF +- xE 0 xFE
gd xL +- xA xFE 0
gd yG +- b 0 yA
gd yH +- b 0 yB
gd yI +- b 0 yC
hxy adj1 0 51965 - 0 0 l th
rect xA yB xE yH
path: M xA yA L xB yB L hc yC L xD yB L xE yA L xF vc L xE yG L xD yH L hc yI L xB yH L xA yG L xL vc Z`,
  mathDivide: `
av adj1 23520
av adj2 5880
av adj3 11760
gd a1 pin 1000 adj1 36745
gd ma1 +- 0 0 a1
gd ma3h +/ 73490 ma1 4
gd ma3w */ 36745 w h
gd maxAdj3 min ma3h ma3w
gd a3 pin 1000 adj3 maxAdj3
gd m4a3 */ -4 a3 1
gd maxAdj2 +- 73490 m4a3 a1
gd a2 pin 0 adj2 maxAdj2
gd dy1 */ h a1 200000
gd yg */ h a2 100000
gd rad */ h a3 100000
gd dx1 */ w 73490 200000
gd y3 +- vc 0 dy1
gd y4 +- vc dy1 0
gd a +- yg rad 0
gd y2 +- y3 0 a
gd y1 +- y2 0 rad
gd y5 +- b 0 y1
gd x1 +- hc 0 dx1
gd x3 +- hc dx1 0
gd x2 +- hc 0 rad
hxy - 0 0 adj1 1000 36745 l y3
hxy - 0 0 adj2 0 maxAdj2 r y2
hxy adj3 1000 maxAdj3 - 0 0 x2 t
rect x1 y3 x3 y4
path: M hc y1 A rad rad 3cd4 21600000 Z M hc y5 A rad rad cd4 21600000 Z M x1 y3 L x3 y3 L x3 y4 L x1 y4 Z`,
  mathEqual: `
av adj1 23520
av adj2 11760
gd a1 pin 0 adj1 36745
gd 2a1 */ a1 2 1
gd mAdj2 +- 100000 0 2a1
gd a2 pin 0 adj2 mAdj2
gd dy1 */ h a1 100000
gd dy2 */ h a2 200000
gd dx1 */ w 73490 200000
gd y2 +- vc 0 dy2
gd y3 +- vc dy2 0
gd y1 +- y2 0 dy1
gd y4 +- y3 dy1 0
gd x1 +- hc 0 dx1
gd x2 +- hc dx1 0
hxy - 0 0 adj1 0 36745 r y1
hxy - 0 0 adj2 0 mAdj2 l y2
rect x1 y1 x2 y4
path: M x1 y1 L x2 y1 L x2 y2 L x1 y2 Z M x1 y3 L x2 y3 L x2 y4 L x1 y4 Z`,
  mathNotEqual: `
av adj1 23520
av adj2 6600000
av adj3 11760
gd a1 pin 0 adj1 50000
gd crAng pin 4200000 adj2 6600000
gd 2a1 */ a1 2 1
gd maxAdj3 +- 100000 0 2a1
gd a3 pin 0 adj3 maxAdj3
gd dy1 */ h a1 100000
gd dy2 */ h a3 200000
gd dx1 */ w 73490 200000
gd x1 +- hc 0 dx1
gd x8 +- hc dx1 0
gd y2 +- vc 0 dy2
gd y3 +- vc dy2 0
gd y1 +- y2 0 dy1
gd y4 +- y3 dy1 0
gd cadj2 +- crAng 0 cd4
gd xadj2 tan hd2 cadj2
gd len mod xadj2 hd2 0
gd bhw */ len dy1 hd2
gd bhw2 */ bhw 1 2
gd x7 +- hc xadj2 bhw2
gd dx67 */ xadj2 y1 hd2
gd x6 +- x7 0 dx67
gd dx57 */ xadj2 y2 hd2
gd x5 +- x7 0 dx57
gd dx47 */ xadj2 y3 hd2
gd x4 +- x7 0 dx47
gd dx37 */ xadj2 y4 hd2
gd x3 +- x7 0 dx37
gd rx7 +- x7 bhw 0
gd rx6 +- x6 bhw 0
gd rx5 +- x5 bhw 0
gd rx4 +- x4 bhw 0
gd rx3 +- x3 bhw 0
gd dx7 */ dy1 hd2 len
gd rxt +- x7 dx7 0
gd lxt +- rx7 0 dx7
gd rx ?: cadj2 rxt rx7
gd lx ?: cadj2 x7 lxt
gd dy3 */ dy1 xadj2 len
gd dy4 +- 0 0 dy3
gd ry ?: cadj2 dy3 t
gd ly ?: cadj2 t dy4
gd dlx +- w 0 rx
gd drx +- w 0 lx
gd dly +- h 0 ry
gd dry +- h 0 ly
hxy - 0 0 adj1 0 50000 r y1
hpolar - 0 0 adj2 4200000 6600000 lx t
hxy - 0 0 adj3 0 maxAdj3 r y2
rect x1 y1 x8 y4
path: M x1 y1 L x6 y1 L lx ly L rx ry L rx6 y1 L x8 y1 L x8 y2 L rx5 y2 L rx4 y3 L x8 y3 L x8 y4 L rx3 y4 L drx dry L dlx dly L x3 y4 L x1 y4 L x1 y3 L x4 y3 L x5 y2 L x1 y2 Z`,

  // ----- flowchart --------------------------------------------------------------------------
  flowChartProcess: `
path w=1 h=1: M 0 0 L 1 0 L 1 1 L 0 1 Z`,
  flowChartAlternateProcess: `
gd x2 +- r 0 ssd6
gd y2 +- b 0 ssd6
gd il */ ssd6 29289 100000
gd ir +- r 0 il
gd ib +- b 0 il
rect il il ir ib
path: M l ssd6 A ssd6 ssd6 cd2 cd4 L x2 t A ssd6 ssd6 3cd4 cd4 L r y2 A ssd6 ssd6 0 cd4 L ssd6 b A ssd6 ssd6 cd4 cd4 Z`,
  flowChartDecision: `
gd ir */ w 3 4
gd ib */ h 3 4
rect wd4 hd4 ir ib
path w=2 h=2: M 0 1 L 1 0 L 2 1 L 1 2 Z`,
  flowChartInputOutput: `
gd x5 */ w 4 5
rect wd5 t x5 b
path w=5 h=5: M 0 5 L 1 0 L 5 0 L 4 5 Z`,
  flowChartPredefinedProcess: `
gd x2 */ w 7 8
rect wd8 t x2 b
path stroke=false w=1 h=1: M 0 0 L 1 0 L 1 1 L 0 1 Z
path fill=none w=8 h=8: M 1 0 L 1 8 M 7 0 L 7 8
path fill=none w=1 h=1: M 0 0 L 1 0 L 1 1 L 0 1 Z`,
  flowChartInternalStorage: `
rect wd8 hd8 r b
path stroke=false w=1 h=1: M 0 0 L 1 0 L 1 1 L 0 1 Z
path fill=none w=8 h=8: M 1 0 L 1 8 M 0 1 L 8 1
path fill=none w=1 h=1: M 0 0 L 1 0 L 1 1 L 0 1 Z`,
  flowChartDocument: `
gd y1 */ h 17322 21600
rect l t r y1
path w=21600 h=21600: M 0 0 L 21600 0 L 21600 17322 C 10800 17322 10800 23922 0 20172 Z`,
  flowChartMultidocument: `
gd y2 */ h 3675 21600
gd y8 */ h 20782 21600
gd x5 */ w 18595 21600
rect l y2 x5 y8
path stroke=false w=21600 h=21600: M 0 20782 C 9298 23542 9298 18022 18595 18022 L 18595 3675 L 0 3675 Z M 1532 3675 L 1532 1815 L 20000 1815 L 20000 16252 C 19298 16252 18595 16352 18595 16352 L 18595 3675 Z M 2972 1815 L 2972 0 L 21600 0 L 21600 14392 C 20800 14392 20000 14467 20000 14467 L 20000 1815 Z
path fill=none w=21600 h=21600: M 0 3675 L 18595 3675 L 18595 18022 C 9298 18022 9298 23542 0 20782 Z M 1532 3675 L 1532 1815 L 20000 1815 L 20000 16252 C 19298 16252 18595 16352 18595 16352 M 2972 1815 L 2972 0 L 21600 0 L 21600 14392 C 20800 14392 20000 14467 20000 14467`,
  flowChartTerminator: `
gd il */ w 1018 21600
gd ir */ w 20582 21600
gd it */ h 3163 21600
gd ib */ h 18437 21600
rect il it ir ib
path w=21600 h=21600: M 3475 0 L 18125 0 A 3475 10800 3cd4 cd2 L 3475 21600 A 3475 10800 cd4 cd2 Z`,
  flowChartPreparation: `
gd x2 */ w 4 5
rect wd5 t x2 b
path w=10 h=10: M 0 5 L 2 0 L 8 0 L 10 5 L 8 10 L 2 10 Z`,
  flowChartManualInput: `
rect l hd5 r b
path w=5 h=5: M 0 1 L 5 0 L 5 5 L 0 5 Z`,
  flowChartManualOperation: `
gd x3 */ w 4 5
rect wd5 t x3 b
path w=5 h=5: M 0 0 L 5 0 L 4 5 L 1 5 Z`,
  flowChartConnector: `${ELLIPSE_RECT}
path: ${ELLIPSE}`,
  flowChartOffpageConnector: `
gd y1 */ h 4 5
rect l t r y1
path w=10 h=10: M 0 0 L 10 0 L 10 8 L 5 10 L 0 8 Z`,
  flowChartPunchedCard: `
rect l hd5 r b
path w=5 h=5: M 0 1 L 1 0 L 5 0 L 5 5 L 0 5 Z`,
  flowChartPunchedTape: `
gd ib */ h 4 5
rect l hd5 r ib
path w=20 h=20: M 0 2 A 5 2 cd2 -10800000 A 5 2 cd2 cd2 L 20 18 A 5 2 0 -10800000 A 5 2 0 cd2 Z`,
  flowChartSummingJunction: `${ELLIPSE_RECT}
path stroke=false: ${ELLIPSE}
path fill=none: M il it L ir ib M ir it L il ib
path fill=none: ${ELLIPSE}`,
  flowChartOr: `${ELLIPSE_RECT}
path stroke=false: ${ELLIPSE}
path fill=none: M hc t L hc b M l vc L r vc
path fill=none: ${ELLIPSE}`,
  flowChartCollate: `
gd ir */ w 3 4
gd ib */ h 3 4
rect wd4 hd4 ir ib
path w=2 h=2: M 0 0 L 2 0 L 1 1 L 2 2 L 0 2 L 1 1 Z`,
  flowChartSort: `
gd ir */ w 3 4
gd ib */ h 3 4
rect wd4 hd4 ir ib
path stroke=false w=2 h=2: M 0 1 L 1 0 L 2 1 L 1 2 Z
path fill=none w=2 h=2: M 0 1 L 2 1
path fill=none w=2 h=2: M 0 1 L 1 0 L 2 1 L 1 2 Z`,
  flowChartExtract: `
gd x2 */ w 3 4
rect wd4 vc x2 b
path w=2 h=2: M 0 2 L 1 0 L 2 2 Z`,
  flowChartMerge: `
gd x2 */ w 3 4
rect wd4 t x2 vc
path w=2 h=2: M 0 0 L 2 0 L 1 2 Z`,
  flowChartOnlineStorage: `
gd x2 */ w 5 6
rect wd6 t x2 b
path w=6 h=6: M 1 0 L 6 0 A 1 3 3cd4 -10800000 L 1 6 A 1 3 cd4 cd2 Z`,
  flowChartDelay: `
gd idx cos wd2 2700000
gd idy sin hd2 2700000
gd ir +- hc idx 0
gd it +- vc 0 idy
gd ib +- vc idy 0
rect l it ir ib
path: M l t L hc t A wd2 hd2 3cd4 cd2 L l b Z`,
  flowChartMagneticTape: `
gd ang1 at2 w h${ELLIPSE_RECT}
path: M hc b A wd2 hd2 cd4 cd4 A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 ang1 L r ib L r b Z`,
  flowChartMagneticDisk: `
gd y3 */ h 5 6
rect l hd3 r y3
path stroke=false w=6 h=6: M 0 1 A 3 1 cd2 cd2 L 6 5 A 3 1 0 cd2 Z
path fill=none w=6 h=6: M 6 1 A 3 1 0 cd2
path fill=none w=6 h=6: M 0 1 A 3 1 cd2 cd2 L 6 5 A 3 1 0 cd2 Z`,
  flowChartMagneticDrum: `
gd x2 */ w 2 3
rect wd6 t x2 b
path stroke=false w=6 h=6: M 1 0 L 5 0 A 1 3 3cd4 cd2 L 1 6 A 1 3 cd4 cd2 Z
path fill=none w=6 h=6: M 5 6 A 1 3 cd4 cd2
path fill=none w=6 h=6: M 1 0 L 5 0 A 1 3 3cd4 cd2 L 1 6 A 1 3 cd4 cd2 Z`,
  flowChartDisplay: `
gd x2 */ w 5 6
rect wd6 t x2 b
path w=6 h=6: M 0 3 L 1 0 L 5 0 A 1 3 3cd4 cd2 L 1 6 Z`,

  // ----- stars and banners ------------------------------------------------------------------
  irregularSeal1: `
gd x5 */ w 4627 21600
gd x12 */ w 8485 21600
gd x21 */ w 16702 21600
gd x24 */ w 14522 21600
gd y3 */ h 6320 21600
gd y6 */ h 8615 21600
gd y9 */ h 13937 21600
gd y18 */ h 13290 21600
rect x5 y3 x21 y9
path w=21600 h=21600: M 10800 5800 L 14522 0 L 14155 5325 L 18380 4457 L 16702 7315 L 21097 8137 L 17607 10475 L 21600 13290 L 16837 12942 L 18145 18095 L 14020 14457 L 13247 19737 L 10532 14935 L 8485 21600 L 7715 15627 L 4762 17617 L 5667 13937 L 135 14587 L 3722 11775 L 0 8615 L 4627 7617 L 370 2295 L 7312 6320 L 8352 2295 Z`,
  irregularSeal2: `
gd x2 */ w 9722 21600
gd x5 */ w 5372 21600
gd x16 */ w 11612 21600
gd x19 */ w 14640 21600
gd y2 */ h 1887 21600
gd y3 */ h 6382 21600
gd y8 */ h 12877 21600
gd y14 */ h 19712 21600
gd y16 */ h 18842 21600
gd y17 */ h 15935 21600
gd y24 */ h 6645 21600
rect x5 y3 x19 y17
path w=21600 h=21600: M 11462 4342 L 14790 0 L 14525 5777 L 18007 3172 L 16380 6532 L 21600 6645 L 16985 9402 L 18270 11290 L 16380 12310 L 18877 15632 L 14640 14350 L 14942 17370 L 12180 15935 L 11612 18842 L 9872 17370 L 8700 19712 L 7527 18125 L 4917 21600 L 4805 18240 L 1285 17825 L 3330 15370 L 0 12877 L 3935 11592 L 1172 8270 L 5372 7817 L 4502 3625 L 8550 6382 L 9722 1887 Z`,
  star4: `
av adj 12500
gd a pin 0 adj 50000
gd iwd2 */ wd2 a 50000
gd ihd2 */ hd2 a 50000
gd sdx cos iwd2 2700000
gd sdy sin ihd2 2700000
gd sx1 +- hc 0 sdx
gd sx2 +- hc sdx 0
gd sy1 +- vc 0 sdy
gd sy2 +- vc sdy 0
gd yAdj +- vc 0 ihd2
hxy - 0 0 adj 0 50000 hc yAdj
rect sx1 sy1 sx2 sy2
path: M l vc L sx1 sy1 L hc t L sx2 sy1 L r vc L sx2 sy2 L hc b L sx1 sy2 Z`,
  star5: `
av adj 19098
av hf 105146
av vf 110557
gd a pin 0 adj 50000
gd swd2 */ wd2 hf 100000
gd shd2 */ hd2 vf 100000
gd svc */ vc vf 100000
gd dx1 cos swd2 1080000
gd dx2 cos swd2 18360000
gd dy1 sin shd2 1080000
gd dy2 sin shd2 18360000
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc dx2 0
gd x4 +- hc dx1 0
gd y1 +- svc 0 dy1
gd y2 +- svc 0 dy2
gd iwd2 */ swd2 a 50000
gd ihd2 */ shd2 a 50000
gd sdx1 cos iwd2 20520000
gd sdx2 cos iwd2 3240000
gd sdy1 sin ihd2 3240000
gd sdy2 sin ihd2 20520000
gd sx1 +- hc 0 sdx1
gd sx2 +- hc 0 sdx2
gd sx3 +- hc sdx2 0
gd sx4 +- hc sdx1 0
gd sy1 +- svc 0 sdy1
gd sy2 +- svc 0 sdy2
gd sy3 +- svc ihd2 0
gd yAdj +- svc 0 ihd2
hxy - 0 0 adj 0 50000 hc yAdj
rect sx1 sy1 sx4 sy3
path: M x1 y1 L sx2 sy1 L hc t L sx3 sy1 L x4 y1 L sx4 sy2 L x3 y2 L hc sy3 L x2 y2 L sx1 sy2 Z`,
  star6: `
av adj 28868
av hf 115470
gd a pin 0 adj 50000
gd swd2 */ wd2 hf 100000
gd dx1 cos swd2 1800000
gd x1 +- hc 0 dx1
gd x2 +- hc dx1 0
gd y2 +- vc hd4 0
gd iwd2 */ swd2 a 50000
gd ihd2 */ hd2 a 50000
gd sdx2 */ iwd2 1 2
gd sx1 +- hc 0 iwd2
gd sx2 +- hc 0 sdx2
gd sx3 +- hc sdx2 0
gd sx4 +- hc iwd2 0
gd sdy1 sin ihd2 3600000
gd sy1 +- vc 0 sdy1
gd sy2 +- vc sdy1 0
gd yAdj +- vc 0 ihd2
hxy - 0 0 adj 0 50000 hc yAdj
rect sx1 sy1 sx4 sy2
path: M x1 hd4 L sx2 sy1 L hc t L sx3 sy1 L x2 hd4 L sx4 vc L x2 y2 L sx3 sy2 L hc b L sx2 sy2 L x1 y2 L sx1 vc Z`,
  star7: `
av adj 34601
av hf 102572
av vf 105210
gd a pin 0 adj 50000
gd swd2 */ wd2 hf 100000
gd shd2 */ hd2 vf 100000
gd svc */ vc vf 100000
gd dx1 */ swd2 97493 100000
gd dx2 */ swd2 78183 100000
gd dx3 */ swd2 43388 100000
gd dy1 */ shd2 62349 100000
gd dy2 */ shd2 22252 100000
gd dy3 */ shd2 90097 100000
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc 0 dx3
gd x4 +- hc dx3 0
gd x5 +- hc dx2 0
gd x6 +- hc dx1 0
gd y1 +- svc 0 dy1
gd y2 +- svc dy2 0
gd y3 +- svc dy3 0
gd iwd2 */ swd2 a 50000
gd ihd2 */ shd2 a 50000
gd sdx1 */ iwd2 97493 100000
gd sdx2 */ iwd2 78183 100000
gd sdx3 */ iwd2 43388 100000
gd sx1 +- hc 0 sdx1
gd sx2 +- hc 0 sdx2
gd sx3 +- hc 0 sdx3
gd sx4 +- hc sdx3 0
gd sx5 +- hc sdx2 0
gd sx6 +- hc sdx1 0
gd sdy1 */ ihd2 90097 100000
gd sdy2 */ ihd2 22252 100000
gd sdy3 */ ihd2 62349 100000
gd sy1 +- svc 0 sdy1
gd sy2 +- svc 0 sdy2
gd sy3 +- svc sdy3 0
gd sy4 +- svc ihd2 0
gd yAdj +- svc 0 ihd2
hxy - 0 0 adj 0 50000 hc yAdj
rect sx2 sy1 sx5 sy3
path: M x1 y2 L sx1 sy2 L x2 y1 L sx3 sy1 L hc t L sx4 sy1 L x5 y1 L sx6 sy2 L x6 y2 L sx5 sy3 L x4 y3 L hc sy4 L x3 y3 L sx2 sy3 Z`,
  star8: `
av adj 37500
gd a pin 0 adj 50000
gd dx1 cos wd2 2700000
gd x1 +- hc 0 dx1
gd x2 +- hc dx1 0
gd dy1 sin hd2 2700000
gd y1 +- vc 0 dy1
gd y2 +- vc dy1 0
gd iwd2 */ wd2 a 50000
gd ihd2 */ hd2 a 50000
gd sdx1 */ iwd2 92388 100000
gd sdx2 */ iwd2 38268 100000
gd sdy1 */ ihd2 92388 100000
gd sdy2 */ ihd2 38268 100000
gd sx1 +- hc 0 sdx1
gd sx2 +- hc 0 sdx2
gd sx3 +- hc sdx2 0
gd sx4 +- hc sdx1 0
gd sy1 +- vc 0 sdy1
gd sy2 +- vc 0 sdy2
gd sy3 +- vc sdy2 0
gd sy4 +- vc sdy1 0
gd yAdj +- vc 0 ihd2
gd idx cos iwd2 2700000
gd idy sin ihd2 2700000
gd il +- hc 0 idx
gd ir +- hc idx 0
gd it +- vc 0 idy
gd ib +- vc idy 0
hxy - 0 0 adj 0 50000 hc yAdj
rect il it ir ib
path: M l vc L sx1 sy2 L x1 y1 L sx2 sy1 L hc t L sx3 sy1 L x2 y1 L sx4 sy2 L r vc L sx4 sy3 L x2 y2 L sx3 sy4 L hc b L sx2 sy4 L x1 y2 L sx1 sy3 Z`,
  star10: `
av adj 42533
av hf 105146
gd a pin 0 adj 50000
gd swd2 */ wd2 hf 100000
gd dx1 */ swd2 95106 100000
gd dx2 */ swd2 58779 100000
gd x1 +- hc 0 dx1
gd x2 +- hc 0 dx2
gd x3 +- hc dx2 0
gd x4 +- hc dx1 0
gd dy1 */ hd2 80902 100000
gd dy2 */ hd2 30902 100000
gd y1 +- vc 0 dy1
gd y2 +- vc 0 dy2
gd y3 +- vc dy2 0
gd y4 +- vc dy1 0
gd iwd2 */ swd2 a 50000
gd ihd2 */ hd2 a 50000
gd sdx1 */ iwd2 80902 100000
gd sdx2 */ iwd2 30902 100000
gd sdy1 */ ihd2 95106 100000
gd sdy2 */ ihd2 58779 100000
gd sx1 +- hc 0 iwd2
gd sx2 +- hc 0 sdx1
gd sx3 +- hc 0 sdx2
gd sx4 +- hc sdx2 0
gd sx5 +- hc sdx1 0
gd sx6 +- hc iwd2 0
gd sy1 +- vc 0 sdy1
gd sy2 +- vc 0 sdy2
gd sy3 +- vc sdy2 0
gd sy4 +- vc sdy1 0
gd yAdj +- vc 0 ihd2
hxy - 0 0 adj 0 50000 hc yAdj
rect sx2 sy2 sx5 sy3
path: M x1 y2 L sx2 sy2 L x2 y1 L sx3 sy1 L hc t L sx4 sy1 L x3 y1 L sx5 sy2 L x4 y2 L sx6 vc L x4 y3 L sx5 sy3 L x3 y4 L sx4 sy4 L hc b L sx3 sy4 L x2 y4 L sx2 sy3 L x1 y3 L sx1 vc Z`,
  star12: `
av adj 37500
gd a pin 0 adj 50000
gd dx1 cos wd2 1800000
gd dy1 sin hd2 3600000
gd x1 +- hc 0 dx1
gd x3 */ w 3 4
gd x4 +- hc dx1 0
gd y1 +- vc 0 dy1
gd y3 */ h 3 4
gd y4 +- vc dy1 0
gd iwd2 */ wd2 a 50000
gd ihd2 */ hd2 a 50000
gd sdx1 cos iwd2 900000
gd sdx2 cos iwd2 2700000
gd sdx3 cos iwd2 4500000
gd sdy1 sin ihd2 4500000
gd sdy2 sin ihd2 2700000
gd sdy3 sin ihd2 900000
gd sx1 +- hc 0 sdx1
gd sx2 +- hc 0 sdx2
gd sx3 +- hc 0 sdx3
gd sx4 +- hc sdx3 0
gd sx5 +- hc sdx2 0
gd sx6 +- hc sdx1 0
gd sy1 +- vc 0 sdy1
gd sy2 +- vc 0 sdy2
gd sy3 +- vc 0 sdy3
gd sy4 +- vc sdy3 0
gd sy5 +- vc sdy2 0
gd sy6 +- vc sdy1 0
gd yAdj +- vc 0 ihd2
hxy - 0 0 adj 0 50000 hc yAdj
rect sx2 sy2 sx5 sy5
path: M l vc L sx1 sy3 L x1 hd4 L sx2 sy2 L wd4 y1 L sx3 sy1 L hc t L sx4 sy1 L x3 y1 L sx5 sy2 L x4 hd4 L sx6 sy3 L r vc L sx6 sy4 L x4 y3 L sx5 sy5 L x3 y4 L sx4 sy6 L hc b L sx3 sy6 L wd4 y4 L sx2 sy5 L x1 y3 L sx1 sy4 Z`,
  star16: star(16, 37500),
  star24: star(24, 37500),
  star32: star(32, 37500),
  ribbon: `
av adj1 16667
av adj2 50000
gd a1 pin 0 adj1 33333
gd a2 pin 25000 adj2 75000
gd x10 +- r 0 wd8
gd dx2 */ w a2 200000
gd x2 +- hc 0 dx2
gd x9 +- hc dx2 0
gd x3 +- x2 wd32 0
gd x8 +- x9 0 wd32
gd x5 +- x2 wd8 0
gd x6 +- x9 0 wd8
gd x4 +- x5 0 wd32
gd x7 +- x6 wd32 0
gd y1 */ h a1 200000
gd y2 */ h a1 100000
gd y4 +- b 0 y2
gd y3 */ y4 1 2
gd hR */ h a1 400000
gd y5 +- b 0 hR
gd y6 +- y2 0 hR
hxy adj2 25000 75000 - 0 0 x2 t
hxy - 0 0 adj1 0 33333 hc y2
rect x2 y2 x9 b
path stroke=false: M l t L x4 t A wd32 hR 3cd4 cd2 L x3 y1 A wd32 hR 3cd4 -10800000 L x8 y2 A wd32 hR cd4 -10800000 L x7 y1 A wd32 hR cd4 cd2 L r t L x10 y3 L r y4 L x9 y4 L x9 y5 A wd32 hR 0 cd4 L x3 b A wd32 hR cd4 cd4 L x2 y4 L l y4 L wd8 y3 Z
path stroke=false fill=darkenLess: M x5 hR A wd32 hR 0 cd4 L x3 y1 A wd32 hR 3cd4 -10800000 L x5 y2 Z M x6 hR A wd32 hR cd2 -5400000 L x8 y1 A wd32 hR 3cd4 cd2 L x6 y2 Z
path fill=none: M l t L x4 t A wd32 hR 3cd4 cd2 L x3 y1 A wd32 hR 3cd4 -10800000 L x8 y2 A wd32 hR cd4 -10800000 L x7 y1 A wd32 hR cd4 cd2 L r t L x10 y3 L r y4 L x9 y4 L x9 y5 A wd32 hR 0 cd4 L x3 b A wd32 hR cd4 cd4 L x2 y4 L l y4 L wd8 y3 Z M x5 hR L x5 y2 M x6 y2 L x6 hR M x2 y4 L x2 y6 M x9 y6 L x9 y4`,
  ribbon2: `
av adj1 16667
av adj2 50000
gd a1 pin 0 adj1 33333
gd a2 pin 25000 adj2 75000
gd x10 +- r 0 wd8
gd dx2 */ w a2 200000
gd x2 +- hc 0 dx2
gd x9 +- hc dx2 0
gd x3 +- x2 wd32 0
gd x8 +- x9 0 wd32
gd x5 +- x2 wd8 0
gd x6 +- x9 0 wd8
gd x4 +- x5 0 wd32
gd x7 +- x6 wd32 0
gd dy1 */ h a1 200000
gd y1 +- b 0 dy1
gd dy2 */ h a1 100000
gd y2 +- b 0 dy2
gd y4 +- t dy2 0
gd y3 +/ y4 b 2
gd hR */ h a1 400000
gd y6 +- b 0 hR
gd y7 +- y1 0 hR
hxy adj2 25000 75000 - 0 0 x2 b
hxy - 0 0 adj1 0 33333 hc y2
rect x2 t x9 y2
path stroke=false: M l b L x4 b A wd32 hR cd4 -10800000 L x3 y1 A wd32 hR cd4 cd2 L x8 y2 A wd32 hR 3cd4 cd2 L x7 y1 A wd32 hR 3cd4 -10800000 L r b L x10 y3 L r y4 L x9 y4 L x9 hR A wd32 hR 0 -5400000 L x3 t A wd32 hR 3cd4 -5400000 L x2 y4 L l y4 L wd8 y3 Z
path stroke=false fill=darkenLess: M x5 y6 A wd32 hR 0 -5400000 L x3 y1 A wd32 hR cd4 cd2 L x5 y2 Z M x6 y6 A wd32 hR cd2 cd4 L x8 y1 A wd32 hR cd4 -10800000 L x6 y2 Z
path fill=none: M l b L wd8 y3 L l y4 L x2 y4 L x2 hR A wd32 hR cd2 cd4 L x8 t A wd32 hR 3cd4 cd4 L x9 y4 L r y4 L x10 y3 L r b L x7 b A wd32 hR cd4 cd2 L x8 y1 A wd32 hR cd4 -10800000 L x3 y2 A wd32 hR 3cd4 -10800000 L x4 y1 A wd32 hR 3cd4 cd2 Z M x5 y2 L x5 y6 M x6 y6 L x6 y2 M x2 y7 L x2 y4 M x9 y4 L x9 y7`,
  ellipseRibbon: `
av adj1 25000
av adj2 50000
av adj3 12500
gd a1 pin 0 adj1 100000
gd a2 pin 25000 adj2 75000
gd q10 +- 100000 0 a1
gd q11 */ q10 1 2
gd q12 +- a1 0 q11
gd minAdj3 max 0 q12
gd a3 pin minAdj3 adj3 a1
gd dx2 */ w a2 200000
gd x2 +- hc 0 dx2
gd x3 +- x2 wd8 0
gd x4 +- r 0 x3
gd x5 +- r 0 x2
gd x6 +- r 0 wd8
gd dy1 */ h a3 100000
gd f1 */ 4 dy1 w
gd q1 */ x3 x3 w
gd q2 +- x3 0 q1
gd y1 */ f1 q2 1
gd cx1 */ x3 1 2
gd cy1 */ f1 cx1 1
gd cx2 +- r 0 cx1
gd q1b */ h a1 100000
gd dy3 +- q1b 0 dy1
gd q3 */ x2 x2 w
gd q4 +- x2 0 q3
gd q5 */ f1 q4 1
gd y3 +- q5 dy3 0
gd q6 +- dy1 dy3 y3
gd q7 +- q6 dy1 0
gd cy3 +- q7 dy3 0
gd rh +- b 0 q1b
gd q8 */ dy1 14 16
gd y2 +/ q8 rh 2
gd y5 +- q5 rh 0
gd y6 +- y3 rh 0
gd cx4 */ x2 1 2
gd q9 */ f1 cx4 1
gd cy4 +- q9 rh 0
gd cx5 +- r 0 cx4
gd cy6 +- cy3 rh 0
gd y7 +- y1 dy3 0
gd cy7 +- q1b q1b y7
gd y8 +- b 0 dy1
hxy adj2 25000 75000 - 0 0 x2 t
hxy - 0 0 adj1 0 100000 hc rh
hxy - 0 0 adj3 minAdj3 a1 l y8
rect x2 y2 x5 rh
path stroke=false: M l t Q cx1 cy1 x3 y1 L x2 y3 Q hc cy3 x5 y3 L x4 y1 Q cx2 cy1 r t L x6 y2 L r rh Q cx5 cy4 x5 y5 L x5 y6 Q hc cy6 x2 y6 L x2 y5 Q cx4 cy4 l rh L wd8 y2 Z
path stroke=false fill=darkenLess: M x3 y7 L x3 y1 L x2 y3 Q hc cy3 x5 y3 L x4 y1 L x4 y7 Q hc cy7 x3 y7 Z
path fill=none: M l t Q cx1 cy1 x3 y1 L x2 y3 Q hc cy3 x5 y3 L x4 y1 Q cx2 cy1 r t L x6 y2 L r rh Q cx5 cy4 x5 y5 L x5 y6 Q hc cy6 x2 y6 L x2 y5 Q cx4 cy4 l rh L wd8 y2 Z M x2 y5 L x2 y3 M x5 y3 L x5 y5 M x3 y1 L x3 y7 M x4 y7 L x4 y1`,
  ellipseRibbon2: `
av adj1 25000
av adj2 50000
av adj3 12500
gd a1 pin 0 adj1 100000
gd a2 pin 25000 adj2 75000
gd q10 +- 100000 0 a1
gd q11 */ q10 1 2
gd q12 +- a1 0 q11
gd minAdj3 max 0 q12
gd a3 pin minAdj3 adj3 a1
gd dx2 */ w a2 200000
gd x2 +- hc 0 dx2
gd x3 +- x2 wd8 0
gd x4 +- r 0 x3
gd x5 +- r 0 x2
gd x6 +- r 0 wd8
gd dy1 */ h a3 100000
gd f1 */ 4 dy1 w
gd q1 */ x3 x3 w
gd q2 +- x3 0 q1
gd u1 */ f1 q2 1
gd y1 +- b 0 u1
gd cx1 */ x3 1 2
gd cu1 */ f1 cx1 1
gd cy1 +- b 0 cu1
gd cx2 +- r 0 cx1
gd q1b */ h a1 100000
gd dy3 +- q1b 0 dy1
gd q3 */ x2 x2 w
gd q4 +- x2 0 q3
gd q5 */ f1 q4 1
gd u3 +- q5 dy3 0
gd y3 +- b 0 u3
gd q6 +- dy1 dy3 u3
gd q7 +- q6 dy1 0
gd cu3 +- q7 dy3 0
gd cy3 +- b 0 cu3
gd rh +- b 0 q1b
gd q8 */ dy1 14 16
gd u2 +/ q8 rh 2
gd y2 +- b 0 u2
gd u5 +- q5 rh 0
gd y5 +- b 0 u5
gd u6 +- u3 rh 0
gd y6 +- b 0 u6
gd cx4 */ x2 1 2
gd q9 */ f1 cx4 1
gd cu4 +- q9 rh 0
gd cy4 +- b 0 cu4
gd cx5 +- r 0 cx4
gd cu6 +- cu3 rh 0
gd cy6 +- b 0 cu6
gd u7 +- u1 dy3 0
gd y7 +- b 0 u7
gd cu7 +- q1b q1b u7
gd cy7 +- b 0 cu7
hxy adj2 25000 75000 - 0 0 x2 b
hxy - 0 0 adj1 0 100000 hc q1b
hxy - 0 0 adj3 minAdj3 a1 l dy1
rect x2 q1b x5 y2
path stroke=false: M l b Q cx1 cy1 x3 y1 L x2 y3 Q hc cy3 x5 y3 L x4 y1 Q cx2 cy1 r b L x6 y2 L r q1b Q cx5 cy4 x5 y5 L x5 y6 Q hc cy6 x2 y6 L x2 y5 Q cx4 cy4 l q1b L wd8 y2 Z
path stroke=false fill=darkenLess: M x3 y7 L x3 y1 L x2 y3 Q hc cy3 x5 y3 L x4 y1 L x4 y7 Q hc cy7 x3 y7 Z
path fill=none: M l b L wd8 y2 L l q1b Q cx4 cy4 x2 y5 L x2 y6 Q hc cy6 x5 y6 L x5 y5 Q cx5 cy4 r q1b L x6 y2 L r b Q cx2 cy1 x4 y1 L x5 y3 Q hc cy3 x2 y3 L x3 y1 Q cx1 cy1 l b Z M x2 y3 L x2 y5 M x5 y5 L x5 y3 M x3 y7 L x3 y1 M x4 y1 L x4 y7`,
  verticalScroll: `
av adj 12500
gd a pin 0 adj 25000
gd ch */ ss a 100000
gd ch2 */ ch 1 2
gd ch4 */ ch 1 4
gd x3 +- ch ch2 0
gd x4 +- ch ch 0
gd x6 +- r 0 ch
gd x7 +- r 0 ch2
gd x5 +- x6 0 ch2
gd y3 +- b 0 ch
gd y4 +- b 0 ch2
hxy adj 0 25000 - 0 0 ch t
rect ch ch x6 y4
path stroke=false: M ch2 b A ch2 ch2 cd4 -5400000 L ch2 y4 A ch4 ch4 cd4 -10800000 L ch y3 L ch ch2 A ch2 ch2 cd2 cd4 L x7 t A ch2 ch2 3cd4 cd2 L x6 ch L x6 y4 A ch2 ch2 0 cd4 Z M x4 ch2 A ch2 ch2 0 cd4 A ch4 ch4 cd4 cd2 Z
path stroke=false fill=darkenLess: M x4 ch2 A ch2 ch2 0 cd4 A ch4 ch4 cd4 cd2 Z M ch y4 A ch2 ch2 0 3cd4 A ch4 ch4 3cd4 cd2 Z
path fill=none: M ch y3 L ch ch2 A ch2 ch2 cd2 cd4 L x7 t A ch2 ch2 3cd4 cd2 L x6 ch L x6 y4 A ch2 ch2 0 cd4 L ch2 b A ch2 ch2 cd4 cd2 Z M x3 t A ch2 ch2 3cd4 cd2 A ch4 ch4 cd4 cd2 L x4 ch2 M x6 ch L x3 ch M ch2 y3 A ch4 ch4 3cd4 cd2 L ch y4 M ch2 b A ch2 ch2 cd4 -5400000 L ch y3`,
  horizontalScroll: `
av adj 12500
gd a pin 0 adj 25000
gd ch */ ss a 100000
gd ch2 */ ch 1 2
gd ch4 */ ch 1 4
gd y3 +- ch ch2 0
gd y4 +- ch ch 0
gd y6 +- b 0 ch
gd y7 +- b 0 ch2
gd y5 +- y6 0 ch2
gd x3 +- r 0 ch
gd x4 +- r 0 ch2
hxy - 0 0 adj 0 25000 l ch
rect ch ch x4 y6
path stroke=false: M r ch2 A ch2 ch2 0 cd4 L x4 ch2 A ch4 ch4 0 cd2 L x3 ch L ch2 ch A ch2 ch2 3cd4 -5400000 L l y7 A ch2 ch2 cd2 -10800000 L ch y6 L x4 y6 A ch2 ch2 cd4 -5400000 Z M ch2 y4 A ch2 ch2 cd4 -5400000 A ch4 ch4 0 -10800000 Z
path stroke=false fill=darkenLess: M ch2 y4 A ch2 ch2 cd4 -5400000 A ch4 ch4 0 -10800000 Z M x4 ch A ch2 ch2 cd4 -16200000 A ch4 ch4 cd2 -10800000 Z
path fill=none: M l y3 A ch2 ch2 cd2 cd4 L x3 ch L x3 ch2 A ch2 ch2 cd2 cd2 L r y5 A ch2 ch2 0 cd4 L ch y6 L ch y7 A ch2 ch2 0 cd2 Z M x3 ch L x4 ch A ch2 ch2 cd4 -5400000 M x4 ch L x4 ch2 A ch4 ch4 0 cd2 M ch2 y4 L ch2 y3 A ch4 ch4 cd2 cd2 A ch2 ch2 0 cd2 M ch y3 L ch y6`,
  wave: `
av adj1 12500
av adj2 0
gd a1 pin 0 adj1 20000
gd a2 pin -10000 adj2 10000
gd y1 */ h a1 100000
gd dy2 */ y1 10 3
gd y2 +- y1 0 dy2
gd y3 +- y1 dy2 0
gd y4 +- b 0 y1
gd y5 +- y4 0 dy2
gd y6 +- y4 dy2 0
gd dx1 */ w a2 100000
gd of2 */ w a2 50000
gd x1 abs dx1
gd dx2 ?: of2 0 of2
gd x2 +- l 0 dx2
gd dx5 ?: of2 of2 0
gd x5 +- r 0 dx5
gd dx3 +/ dx2 x5 3
gd x3 +- x2 dx3 0
gd x4 +/ x3 x5 2
gd x6 +- l dx5 0
gd x10 +- r dx2 0
gd x7 +- x6 dx3 0
gd x8 +/ x7 x10 2
gd xAdj +- hc dx1 0
gd il max x2 x6
gd ir min x5 x10
gd it */ h a1 50000
gd ib +- b 0 it
hxy - 0 0 adj1 0 20000 l y1
hxy adj2 -10000 10000 - 0 0 xAdj b
rect il it ir ib
path: M x2 y1 C x3 y2 x4 y3 x5 y1 L x10 y4 C x8 y6 x7 y5 x6 y4 Z`,
  doubleWave: `
av adj1 6250
av adj2 0
gd a1 pin 0 adj1 12500
gd a2 pin -10000 adj2 10000
gd y1 */ h a1 100000
gd dy2 */ y1 10 3
gd y2 +- y1 0 dy2
gd y3 +- y1 dy2 0
gd y4 +- b 0 y1
gd y5 +- y4 0 dy2
gd y6 +- y4 dy2 0
gd dx1 */ w a2 100000
gd of2 */ w a2 50000
gd x1 abs dx1
gd dx2 ?: of2 0 of2
gd x2 +- l 0 dx2
gd dx8 ?: of2 of2 0
gd x8 +- r 0 dx8
gd dx3 +/ dx2 x8 6
gd x3 +- x2 dx3 0
gd dx4 +/ dx2 x8 3
gd x4 +- x2 dx4 0
gd x5 +/ x2 x8 2
gd x6 +- x5 dx3 0
gd x7 +/ x6 x8 2
gd x9 +- l dx8 0
gd x15 +- r dx2 0
gd x10 +- x9 dx3 0
gd x11 +- x9 dx4 0
gd x12 +/ x9 x15 2
gd x13 +- x12 dx3 0
gd x14 +/ x13 x15 2
gd xAdj +- hc dx1 0
gd il max x2 x9
gd ir min x8 x15
gd it */ h a1 50000
gd ib +- b 0 it
hxy - 0 0 adj1 0 12500 l y1
hxy adj2 -10000 10000 - 0 0 xAdj b
rect il it ir ib
path: M x2 y1 C x3 y2 x4 y3 x5 y1 C x6 y2 x7 y3 x8 y1 L x15 y4 C x14 y6 x13 y5 x12 y4 C x11 y6 x10 y5 x9 y4 Z`,

  // ----- callouts ---------------------------------------------------------------------------
  wedgeRectCallout: `
av adj1 -20833
av adj2 62500
gd dxPos */ w adj1 100000
gd dyPos */ h adj2 100000
gd xPos +- hc dxPos 0
gd yPos +- vc dyPos 0
gd dq */ dxPos h w
gd ady abs dyPos
gd adq abs dq
gd dz +- ady 0 adq
gd xg1 ?: dxPos 7 2
gd xg2 ?: dxPos 10 5
gd x1 */ w xg1 12
gd x2 */ w xg2 12
gd yg1 ?: dyPos 7 2
gd yg2 ?: dyPos 10 5
gd y1 */ h yg1 12
gd y2 */ h yg2 12
gd t1 ?: dxPos l xPos
gd xl ?: dz l t1
gd t2 ?: dyPos x1 xPos
gd xt ?: dz t2 x1
gd t3 ?: dxPos xPos r
gd xr ?: dz r t3
gd t4 ?: dyPos xPos x1
gd xb ?: dz t4 x1
gd t5 ?: dxPos y1 yPos
gd yl ?: dz y1 t5
gd t6 ?: dyPos t yPos
gd yt ?: dz t6 t
gd t7 ?: dxPos yPos y1
gd yr ?: dz y1 t7
gd t8 ?: dyPos yPos b
gd yb ?: dz t8 b
hxy adj1 -2147483647 2147483647 adj2 -2147483647 2147483647 xPos yPos
rect l t r b
path: M l t L x1 t L xt yt L x2 t L r t L r y1 L xr yr L r y2 L r b L x2 b L xb yb L x1 b L l b L l y2 L xl yl L l y1 Z`,
  wedgeRoundRectCallout: `
av adj1 -20833
av adj2 62500
av adj3 16667
gd dxPos */ w adj1 100000
gd dyPos */ h adj2 100000
gd xPos +- hc dxPos 0
gd yPos +- vc dyPos 0
gd dq */ dxPos h w
gd ady abs dyPos
gd adq abs dq
gd dz +- ady 0 adq
gd xg1 ?: dxPos 7 2
gd xg2 ?: dxPos 10 5
gd x1 */ w xg1 12
gd x2 */ w xg2 12
gd yg1 ?: dyPos 7 2
gd yg2 ?: dyPos 10 5
gd y1 */ h yg1 12
gd y2 */ h yg2 12
gd t1 ?: dxPos l xPos
gd xl ?: dz l t1
gd t2 ?: dyPos x1 xPos
gd xt ?: dz t2 x1
gd t3 ?: dxPos xPos r
gd xr ?: dz r t3
gd t4 ?: dyPos xPos x1
gd xb ?: dz t4 x1
gd t5 ?: dxPos y1 yPos
gd yl ?: dz y1 t5
gd t6 ?: dyPos t yPos
gd yt ?: dz t6 t
gd t7 ?: dxPos yPos y1
gd yr ?: dz y1 t7
gd t8 ?: dyPos yPos b
gd yb ?: dz t8 b
gd u1 */ ss adj3 100000
gd u2 +- r 0 u1
gd v2 +- b 0 u1
gd il */ u1 29289 100000
gd ir +- r 0 il
gd ib +- b 0 il
hxy adj1 -2147483647 2147483647 adj2 -2147483647 2147483647 xPos yPos
rect il il ir ib
path: M l u1 A u1 u1 cd2 cd4 L x1 t L xt yt L x2 t L u2 t A u1 u1 3cd4 cd4 L r y1 L xr yr L r y2 L r v2 A u1 u1 0 cd4 L x2 b L xb yb L x1 b L u1 b A u1 u1 cd4 cd4 L l y2 L xl yl L l y1 Z`,
  wedgeEllipseCallout: `
av adj1 -20833
av adj2 62500
gd dxPos */ w adj1 100000
gd dyPos */ h adj2 100000
gd xPos +- hc dxPos 0
gd yPos +- vc dyPos 0
gd sdx */ dxPos h 1
gd sdy */ dyPos w 1
gd pang at2 sdx sdy
gd stAng +- pang 660000 0
gd enAng +- pang 0 660000
gd dx1 cos wd2 stAng
gd dy1 sin hd2 stAng
gd x1 +- hc dx1 0
gd y1 +- vc dy1 0
gd dx2 cos wd2 enAng
gd dy2 sin hd2 enAng
gd stAng1 at2 dx1 dy1
gd enAng1 at2 dx2 dy2
gd swAng1 +- enAng1 0 stAng1
gd swAng2 +- swAng1 21600000 0
gd swAng ?: swAng1 swAng1 swAng2${ELLIPSE_RECT}
hxy adj1 -2147483647 2147483647 adj2 -2147483647 2147483647 xPos yPos
path: M xPos yPos L x1 y1 A wd2 hd2 stAng1 swAng Z`,
  cloudCallout: `
av adj1 -20833
av adj2 62500
gd dxPos */ w adj1 100000
gd dyPos */ h adj2 100000
gd xPos +- hc dxPos 0
gd yPos +- vc dyPos 0
gd ht cat2 hd2 dxPos dyPos
gd wt sat2 wd2 dxPos dyPos
gd g2 cat2 wd2 ht wt
gd g3 sat2 hd2 ht wt
gd g4 +- hc g2 0
gd g5 +- vc g3 0
gd g6 +- g4 0 xPos
gd g7 +- g5 0 yPos
gd g8 mod g6 g7 0
gd g9 */ ss 6600 21600
gd g10 +- g8 0 g9
gd g11 */ g10 1 3
gd g12 */ ss 1800 21600
gd g13 +- g11 g12 0
gd g14 */ g13 g6 g8
gd g15 */ g13 g7 g8
gd g16 +- g14 xPos 0
gd g17 +- g15 yPos 0
gd g18 */ ss 4800 21600
gd g19 */ g11 2 1
gd g20 +- g18 g19 0
gd g21 */ g20 g6 g8
gd g22 */ g20 g7 g8
gd g23 +- g21 xPos 0
gd g24 +- g22 yPos 0
gd g25 */ ss 1200 21600
gd g26 */ ss 600 21600
gd x23 +- xPos g26 0
gd x24 +- g16 g25 0
gd x25 +- g23 g12 0
gd il */ w 2977 21600
gd it */ h 3262 21600
gd ir */ w 17087 21600
gd ib */ h 17337 21600
hxy adj1 -2147483647 2147483647 adj2 -2147483647 2147483647 xPos yPos
rect il it ir ib
path w=43200 h=43200: ${CLOUD}
path: M x23 yPos A g26 g26 0 21600000 Z
path: M x24 g17 A g25 g25 0 21600000 Z
path: M x25 g24 A g12 g12 0 21600000 Z
path w=43200 h=43200 fill=none: ${CLOUD_ARCS}`,
  callout1: callout(1, false, false),
  callout2: callout(2, false, false),
  callout3: callout(3, false, false),
  accentCallout1: callout(1, false, true),
  accentCallout2: callout(2, false, true),
  accentCallout3: callout(3, false, true),
  borderCallout1: callout(1, true, false),
  borderCallout2: callout(2, true, false),
  borderCallout3: callout(3, true, false),
  accentBorderCallout1: callout(1, true, true),
  accentBorderCallout2: callout(2, true, true),
  accentBorderCallout3: callout(3, true, true),
};
