const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Carnival Network Engineering';
pptx.company = 'Carnival Corporation';
pptx.subject = 'Merged C9200CX competitive comparison';
pptx.title = 'C9200CX Competitive Comparison Master Deck';
pptx.lang = 'en-US';

const C = {
  blue: '244C7C',
  blue2: '1B3F6B',
  text: '2D2D2D',
  dark: '1F2937',
  line: 'BFC7D1',
  gray: 'E8E8EA',
  gray2: 'F5F6F8',
  white: 'FFFFFF',
  green: '1FA74A',
  amber: 'F5A623',
  red: 'D7261E',
};

const FONT = 'Aptos';
const TOTAL = 8;

function hdr(slide, title, subtitle, page) {
  slide.background = { color: C.white };
  slide.addText(title, { x: 0.72, y: 0.48, w: 10.8, h: 0.55, fontFace: FONT, fontSize: 28, color: C.blue, bold: false });
  if (subtitle) slide.addText(subtitle, { x: 0.72, y: 1.0, w: 10.8, h: 0.38, fontFace: FONT, fontSize: 20, color: C.blue, bold: false });
  slide.addText('© 2026 OpenCode / Carnival internal working draft', { x: 0.75, y: 7.0, w: 4.8, h: 0.2, fontFace: FONT, fontSize: 8, color: C.blue2 });
  slide.addText('Competitive Comparison', { x: 10.4, y: 7.0, w: 2.0, h: 0.2, fontFace: FONT, fontSize: 8, color: C.blue2, align: 'right' });
  slide.addText(String(page), { x: 12.45, y: 7.0, w: 0.2, h: 0.2, fontFace: FONT, fontSize: 8, color: C.blue2, align: 'right' });
}

function addTable(slide, rows, opts = {}) {
  slide.addTable(rows, {
    x: opts.x ?? 0.6,
    y: opts.y ?? 1.6,
    w: opts.w ?? 12.1,
    colW: opts.colW,
    rowH: opts.rowH ?? 0.46,
    fontFace: FONT,
    fontSize: opts.fontSize ?? 12,
    color: C.text,
    border: { pt: 0.5, color: 'D9D9DD' },
    margin: 0.05,
    valign: 'mid',
    autoPage: false,
  });
}

function h(text) {
  return { text, options: { bold: true, color: C.dark, fill: { color: C.white }, align: 'center' } };
}

function label(text) {
  return { text, options: { bold: true, fill: { color: C.white }, color: C.dark } };
}

function cell(text, fill = C.gray2, extra = {}) {
  return { text, options: { fill: { color: fill }, color: C.text, align: 'center', ...extra } };
}

function winner(text) {
  return { text: `🏆 ${text}`, options: { fill: { color: C.gray2 }, color: C.green, bold: true, align: 'center' } };
}

function good(text) {
  return { text, options: { fill: { color: C.gray2 }, color: C.green, bold: true, align: 'center' } };
}

function warn(text) {
  return { text, options: { fill: { color: C.gray2 }, color: C.amber, bold: true, align: 'center' } };
}

function bad(text) {
  return { text, options: { fill: { color: C.gray2 }, color: C.red, bold: true, align: 'center' } };
}

function bullets(slide, items, x, y, w, fs = 15) {
  let yy = y;
  for (const item of items) {
    slide.addText([{ text: '• ', options: { color: C.text } }, { text: item, options: { color: C.text } }], {
      x, y: yy, w, h: 0.5, fontFace: FONT, fontSize: fs, breakLine: false, margin: 0
    });
    yy += 0.72;
  }
}

{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  s.addText('Competitive Overview', { x: 0.72, y: 0.44, w: 6.2, h: 0.45, fontFace: FONT, fontSize: 16, color: C.dark });
  s.addText('C9200CX Competitive Comparison', { x: 0.72, y: 0.92, w: 8.8, h: 0.75, fontFace: FONT, fontSize: 28, color: C.blue, bold: false });
  s.addText('Merged from validated Aruba, Arista, and IE3500H comparison decks', { x: 0.72, y: 1.64, w: 7.8, h: 0.35, fontFace: FONT, fontSize: 17, color: C.text });
  s.addText('Prepared for: Carnival', { x: 0.72, y: 2.08, w: 4, h: 0.3, fontFace: FONT, fontSize: 15, color: C.text });
  s.addText('March 2026', { x: 0.72, y: 2.42, w: 2, h: 0.3, fontFace: FONT, fontSize: 15, color: C.text });
  s.addShape(pptx.ShapeType.line, { x: 0.72, y: 2.95, w: 11.8, h: 0, line: { color: C.line, pt: 1 } });
  bullets(s, [
    'Input decks merged: Aruba CX 6000/6100, Arista CCS-710P-12, and Cisco IE-3500H-12P2MU2XE.',
    'Reference model remains Cisco C9200CX-12P-2X2G-E across all tables.',
    'Winners are highlighted per category with green trophy markers.'
  ], 0.92, 3.45, 11.0, 18);
  s.addText('© 2026 OpenCode / Carnival internal working draft', { x: 0.75, y: 7.0, w: 4.8, h: 0.2, fontFace: FONT, fontSize: 8, color: C.blue2 });
  s.addText('Competitive Comparison', { x: 10.4, y: 7.0, w: 2.0, h: 0.2, fontFace: FONT, fontSize: 8, color: C.blue2, align: 'right' });
  s.addText('1', { x: 12.45, y: 7.0, w: 0.2, h: 0.2, fontFace: FONT, fontSize: 8, color: C.blue2, align: 'right' });
}

{
  const s = pptx.addSlide();
  hdr(s, 'Key Technical Findings:', '', 2);
  bullets(s, [
    'Arista CCS-710P-12 is the closest true 1:1 compact competitor to C9200CX: same 12-port fanless PoE form factor with 2x10G SFP+ uplinks.',
    'Aruba CX 6100 is the lightest practical alternative with 10G-capable uplinks and VSF stacking, but it gives up major PoE, security, and telemetry features.',
    'Cisco IE-3500H-12P2MU2XE is not a campus peer; it is the ruggedized winner for harsh environments, with DC power, IP66/IP67, M12, and industrial hardening.',
    'Across feature depth, Cisco C9200CX remains the broadest all-around winner for campus access, while each competitor wins a narrower category.'
  ], 0.92, 1.6, 11.2, 17);
}

{
  const s = pptx.addSlide();
  hdr(s, 'Cisco Catalyst 9200CX Switches', 'Competitive Landscape', 3);
  addTable(s, [
    [h('Platform'), h('Compared Model'), h('Positioning'), h('Key Differentiator')],
    [label('Cisco'), cell('C9200CX-12P-2X2G-E', C.gray), cell('Reference compact campus switch', C.gray), winner('Best all-around campus feature set')],
    [label('Aruba'), cell('CX 6100 (JL679A)', C.white), cell('Closest lightweight Aruba 10G option', C.white), good('Winner: stacking / weight')],
    [label('Aruba'), cell('CX 6000 (R8N89A)', C.gray), cell('Lower-end 1G-uplink option', C.gray), bad('Lacks 10G uplinks')],
    [label('Arista'), cell('CCS-710P-12', C.white), cell('Closest 1:1 compact Arista match', C.white), good('Winner: closest clean near-match')],
    [label('Cisco Industrial'), cell('IE-3500H-12P2MU2XE', C.gray), cell('Industrial / harsh-environment near-match', C.gray), good('Winner: rugged / IP67 / DC')],
  ], { y: 1.7, colW: [1.6, 3.4, 3.7, 3.4], fontSize: 13, rowH: 0.58 });
}

{
  const s = pptx.addSlide();
  hdr(s, 'Catalyst 9200CX Hardware Comparison', '', 4);
  addTable(s, [
    [h('Category'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('Weight'), cell('6.6 lb', C.gray), winner('6.13 lb'), good('6.17 lb'), cell('not explicit in final deck', C.gray), bad('11.40 lb')],
    [label('10G Uplinks'), winner('2x10G SFP+ + 2x1G copper'), good('2x1/10G SFP+ + 2x1G'), bad('No'), good('2x10G SFP+'), good('2x10G/1G SFP')],
    [label('PoE Budget'), winner('240W'), bad('139W'), bad('139W'), warn('234W'), winner('240W')],
    [label('Form Factor'), good('Compact fanless 1U'), good('Compact fanless 1U'), good('Compact fanless 1U'), good('Compact fanless 1U'), warn('Industrial wall-mount heavy duty')],
    [label('Power Input'), cell('315W AC internal', C.gray), cell('Internal 165W', C.gray), cell('Internal', C.gray), warn('External 150W/280W AC'), winner('12-54V DC industrial')],
    [label('Packet Buffer'), cell('6 MB', C.gray), winner('12.38 MB'), cell('not specified', C.gray), bad('2 MB'), cell('not officially published', C.gray)],
    [label('Stacking'), bad('No'), winner('VSF up to 8'), warn('Not confirmed'), bad('No evidence in deck'), bad('No')],
  ], { y: 1.7, colW: [1.75, 2.1, 2.05, 2.0, 2.0, 2.0], fontSize: 11.5, rowH: 0.56 });
}

{
  const s = pptx.addSlide();
  hdr(s, 'Catalyst 9200CX Hardware Comparison', 'Security / Operations / Environment', 5);
  addTable(s, [
    [h('Category'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('MACsec'), winner('AES-256'), bad('No'), bad('No'), warn('Not fully substantiated'), warn('MACsec-128 on -E')],
    [label('NetFlow / Visibility'), winner('Flexible NetFlow + telemetry'), bad('sFlow only'), bad('sFlow only'), warn('sFlow / CloudVision telemetry'), cell('DNA Center support', C.gray)],
    [label('App Visibility'), winner('NBAR2 / AVC'), bad('No'), bad('No'), bad('No NBAR equivalent'), bad('OT focus, not campus app vis')],
    [label('Campus Fabric'), winner('SD-Access support'), bad('No'), bad('No'), warn('VXLAN/EVPN positioning'), warn('Industrial OT focus')],
    [label('Operating Temp'), cell('-5°C to 45°C', C.gray), cell('0°C to 45°C', C.gray), cell('0°C to 45°C', C.gray), cell('not emphasized as harsh-env', C.gray), winner('-40°C to 75°C')],
    [label('Ingress / Ruggedness'), bad('No IP rating'), bad('Commercial campus'), bad('Commercial campus'), bad('Compact campus edge'), winner('IP66 / IP67')],
  ], { y: 1.7, colW: [1.75, 2.1, 2.05, 2.0, 2.0, 2.0], fontSize: 11.3, rowH: 0.56 });
}

{
  const s = pptx.addSlide();
  hdr(s, 'Catalyst 9200CX Scale Comparison', '', 6);
  addTable(s, [
    [h('Scale Metric'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('MAC'), winner('32K'), bad('8K'), bad('8K'), warn('Not clearly published'), cell('24K', C.gray)],
    [label('IPv4 / Host Routes'), cell('14K total / 4K routing entries', C.gray), bad('1K class'), bad('1K class'), warn('Not explicitly published in source set'), winner('7K IPv4')],
    [label('Forwarding Rate'), cell('50.59 Mpps', C.gray), cell('45.1 Mpps', C.gray), bad('23.8 Mpps'), winner('95 Mpps'), winner('Line rate for all ports')],
    [label('Switching Capacity'), winner('68 Gbps'), winner('68 Gbps'), bad('32 Gbps'), cell('64 Gbps', C.gray), warn('Not officially published for exact SKU')],
    [label('Memory'), cell('4 GB DRAM / 8 GB flash', C.gray), cell('4 GB / 16 GB', C.gray), cell('4 GB / 16 GB corrected', C.gray), cell('4 GB / 8 GB', C.gray), winner('8 GB / 5.1 GB user storage')],
  ], { y: 1.75, colW: [2.0, 2.0, 2.0, 2.0, 2.0, 2.0], fontSize: 11.2, rowH: 0.6 });
  s.addShape(pptx.ShapeType.rect, { x: 0.6, y: 5.95, w: 12.1, h: 0.38, fill: { color: C.blue }, line: { color: C.blue, pt: 0 } });
  s.addText('C9200CX wins campus-scale balance; IE3500H wins rugged route scale; Arista wins compact forwarding efficiency', { x: 0.8, y: 6.03, w: 11.7, h: 0.2, fontFace: FONT, fontSize: 13, color: C.white, bold: true, align: 'center' });
}

{
  const s = pptx.addSlide();
  hdr(s, 'Catalyst 9200CX Competitive Summary', '', 7);
  addTable(s, [
    [h('Category'), h('Winner'), h('Reason')],
    [label('Closest 1:1 equivalent'), winner('Arista CCS-710P-12'), cell('Same compact fanless 12-port PoE role with 2x10G SFP+ uplinks.', C.gray)],
    [label('Best all-around campus features'), winner('Cisco C9200CX'), cell('PoE headroom, MACsec-256, NetFlow, app visibility, SD-Access, richer published L3/telemetry.', C.white)],
    [label('Best rugged / marine / OT fit'), winner('Cisco IE-3500H-12P2MU2XE'), cell('IP66/IP67, DC power, M12, -40°C to 75°C, industrial hardening.', C.gray)],
    [label('Best weight / stacking trade-off'), winner('Aruba CX 6100'), cell('Lowest validated weight and VSF stacking, but with major feature trade-offs.', C.white)],
    [label('Lowest-capability option'), bad('Aruba CX 6000'), cell('Viable only if 1G uplinks and basic L2 access are acceptable.', C.gray)],
  ], { y: 1.75, colW: [2.8, 3.0, 6.3], fontSize: 12.5, rowH: 0.62 });
}

{
  const s = pptx.addSlide();
  hdr(s, 'What We Need from the BU & PM', '', 8);
  bullets(s, [
    'Authoritative C9200CX MTBF and environmental guidance to compare fairly against rugged and campus competitors.',
    'Explicit public Cisco guidance for marine / hospitality deployments where weight, humidity, and non-standard mounting matter.',
    'A customer-ready positioning narrative for why C9200CX should be chosen over Aruba / Arista / IE3500 when requirements are mixed.',
    'Clear field guidance on when customers should move to IE3500H instead of forcing C9200CX into harsh-environment roles.',
    'Published competitive positioning on Arista 710P, since it is the cleanest compact non-Cisco near-match in this study.'
  ], 0.92, 1.7, 11.25, 17);
}

pptx.writeFile({ fileName: 'master_c9200cx_competitive_comparison.pptx' });
