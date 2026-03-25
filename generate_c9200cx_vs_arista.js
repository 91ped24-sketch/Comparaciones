const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Carnival Network Engineering';
pptx.company = 'Carnival Corporation';
pptx.subject = 'Cisco C9200CX vs Arista 710P - Corrected & Validated';
pptx.title = 'Cisco C9200CX-12P-2X2G-E vs Arista CCS-710P-12';
pptx.lang = 'en-US';

const C = {
  navy: '0B1F3F',
  dBlue: '14325A',
  blue: '1B5FAA',
  green: '16824A',
  greenBg: 'E8F5EE',
  amber: 'B47D1B',
  amberBg: 'FEF7E8',
  red: 'B91C1C',
  redBg: 'FEF2F2',
  text: '1E293B',
  sub: '475569',
  mut: '94A3B8',
  line: 'CBD5E1',
  row1: 'F8FAFC',
  row2: 'FFFFFF',
  hdr: '0B1F3F',
  hdrTx: 'FFFFFF',
  softBg: 'F1F5F9',
};

const FONT = 'Calibri';
const TOTAL = 13;

function hdr(t) { return { text: t, options: { bold: true, color: C.hdrTx, fill: { color: C.hdr }, align: 'left', valign: 'middle' } }; }
function cell(t, opts) { return { text: t, options: { valign: 'middle', ...opts } }; }
function greenCell(t) { return cell(t, { color: C.green, bold: true }); }
function redCell(t) { return cell(t, { color: C.red, bold: true }); }
function amberCell(t) { return cell(t, { color: C.amber, bold: true }); }

function stripe(rows) {
  return rows.map((r, i) => {
    if (i === 0) return r;
    const bg = i % 2 === 1 ? C.row1 : C.row2;
    return r.map(c => {
      if (typeof c === 'string') return { text: c, options: { fill: { color: bg } } };
      c.options = c.options || {};
      if (!c.options.fill) c.options.fill = { color: bg };
      return c;
    });
  });
}

function addSlideTitle(slide, title, subtitle) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 1.15, fill: { color: C.navy } });
  slide.addText(title, { x: 0.6, y: 0.15, w: 11.8, h: 0.55, fontFace: FONT, fontSize: 24, bold: true, color: C.hdrTx });
  if (subtitle) slide.addText(subtitle, { x: 0.6, y: 0.72, w: 11.5, h: 0.3, fontFace: FONT, fontSize: 11, color: C.mut, italic: true });
}

function addFooter(slide, pg, total) {
  slide.addText(`Carnival | C9200CX vs Arista 710P | Slide ${pg} of ${total}`, {
    x: 0.6, y: 7.05, w: 8.2, h: 0.25, fontFace: FONT, fontSize: 8, color: C.mut
  });
  slide.addText('Validated March 2026', {
    x: 10.5, y: 7.05, w: 2.3, h: 0.25, fontFace: FONT, fontSize: 8, color: C.mut, align: 'right'
  });
  slide.addShape(pptx.ShapeType.line, { x: 0.6, y: 7.02, w: 12.1, h: 0, line: { color: C.line, pt: 0.5 } });
}

function addTable(slide, rows, opts) {
  slide.addTable(stripe(rows), {
    x: opts.x ?? 0.5, y: opts.y ?? 1.4, w: opts.w ?? 12.3,
    border: { type: 'solid', pt: 0.5, color: C.line },
    color: C.text, fontFace: FONT, fontSize: opts.fs ?? 14,
    rowH: opts.rowH ?? 0.42, colW: opts.colW, margin: [0.04, 0.08, 0.04, 0.08],
    valign: 'middle', autoPage: false,
  });
}

function addConclusions(slide, items, y) {
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.5, y, w: 12.3, h: items.length * 0.38 + 0.2, rectRadius: 0.04, fill: { color: C.softBg }, line: { color: C.line, pt: 1 } });
  items.forEach((txt, i) => {
    slide.addText([{ text: '▶  ', options: { bold: true, color: C.blue } }, { text: txt, options: { color: C.text } }], {
      x: 0.7, y: y + 0.1 + i * 0.38, w: 11.9, h: 0.36, fontFace: FONT, fontSize: 14
    });
  });
}

{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  s.addText('Cisco C9200CX-12P-2X2G-E\nvs. Arista CCS-710P-12', {
    x: 0.8, y: 0.8, w: 11.8, h: 2.0, fontFace: FONT, fontSize: 34, bold: true, color: C.hdrTx, lineSpacingMultiple: 1.2
  });
  s.addShape(pptx.ShapeType.line, { x: 0.8, y: 3.0, w: 4.8, h: 0, line: { color: C.blue, pt: 3 } });
  s.addText('Compact Fanless 12-Port PoE Access Switch Comparison', {
    x: 0.8, y: 3.3, w: 11.6, h: 0.6, fontFace: FONT, fontSize: 22, color: C.mut
  });
  s.addText('Context: Same-format validated comparison using official Cisco and Arista sources\nPrepared For: Carnival', {
    x: 0.8, y: 4.2, w: 9.0, h: 0.9, fontFace: FONT, fontSize: 14, color: C.mut, lineSpacingMultiple: 1.4
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.6, w: 6.2, h: 1.0, rectRadius: 0.06, fill: { color: C.dBlue }, line: { color: C.blue, pt: 1 } });
  s.addText('Validated against official vendor product pages and datasheets\nMarch 2026', {
    x: 1.0, y: 5.8, w: 5.8, h: 0.6, fontFace: FONT, fontSize: 13, color: C.hdrTx, bold: true, align: 'center'
  });
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Executive Summary', 'Arista CCS-710P-12 is the closest clean 1:1 near-match found');
  addTable(s, [
    [hdr('Metric'), hdr('Cisco C9200CX-12P-2X2G-E'), hdr('Arista CCS-710P-12')],
    ['Downlink Ports', '12x 1G PoE+', '12x 1G PoE+'],
    ['Uplinks', '2x 10G SFP+ + 2x 1G copper', greenCell('2x 10G SFP+')],
    ['PoE Budget', greenCell('240W'), '234W'],
    ['Switching Capacity', greenCell('68 Gbps'), '64 Gbps'],
    ['Forwarding Rate', '50.59 Mpps', greenCell('95 Mpps')],
    ['Form Factor', greenCell('Compact, fanless, 1RU'), greenCell('Compact, fanless, 1RU')],
    ['Power Input', '315W AC internal', amberCell('150W / 280W external AC PSU options')],
    ['Management OS', 'Cisco IOS XE', 'Arista EOS'],
    ['Telemetry', 'Model-driven telemetry, NetFlow', 'CloudVision telemetry, sFlow'],
  ], { y: 1.35, colW: [3.0, 4.5, 4.8] });
  addConclusions(s, [
    'This is a much cleaner 1:1 comparison than IE3500: both are compact, fanless, 12-port PoE campus/edge switches with 10G uplinks.',
    'Cisco leads slightly on PoE budget and uplink flexibility; Arista leads on published throughput efficiency and power draw.'
  ], 5.6);
  addFooter(s, 2, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Why CCS-710P-12 Is the Correct Arista Equivalent');
  addTable(s, [
    [hdr('Selection Test'), hdr('Result'), hdr('Reason')],
    ['12-port compact PoE intent', greenCell('Match'), 'Arista 710P-12 is a compact, fanless 12-port PoE campus switch.'],
    ['10G uplinks', greenCell('Match'), 'Arista provides 2x 10G SFP+ uplinks, matching the core uplink intent.'],
    ['Branch / access deployment', greenCell('Match'), 'Both target space-constrained access-layer deployments.'],
    ['Fanless / quiet environments', greenCell('Match'), 'Both are designed for silent or quiet installation areas.'],
    ['Exact uplink parity', amberCell('Partial'), 'Cisco adds 2x 1G copper uplinks; Arista uses SFP+ uplinks only.'],
  ], { y: 1.45, colW: [3.7, 2.0, 6.4] });
  addConclusions(s, [
    'CCS-710P-12 is the best Arista near-match because it aligns on form factor, PoE access role, and compact deployment assumptions.',
    'Unlike the IE3500 comparison, no product-class mismatch is required here.'
  ], 4.9);
  addFooter(s, 3, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Ports, Uplinks, and Physical Interfaces');
  addTable(s, [
    [hdr('Parameter'), hdr('Cisco C9200CX-12P-2X2G-E'), hdr('Arista CCS-710P-12')],
    ['Primary Ports', '12x 10/100/1000 RJ-45 PoE+', '12x 10/100/1000 RJ-45 PoE+'],
    ['10G Uplinks', '2x 10G SFP+', '2x 10G SFP+'],
    ['Additional Copper Uplinks', greenCell('2x 1G copper'), redCell('None')],
    ['Console / USB', 'Console + USB management ports', 'Console + USB'],
    ['Mounting', 'Rack / wall / table', greenCell('Rack / wall / ceiling / magnetic options')],
    ['Acoustics', greenCell('Fanless / silent'), greenCell('Fanless / silent')],
  ], { y: 1.35, colW: [3.1, 4.5, 4.7] });
  addConclusions(s, [
    'Cisco’s only meaningful physical advantage is the extra 2x 1G copper uplinks.',
    'Arista counters with broader mounting flexibility for compact remote deployments.'
  ], 5.4);
  addFooter(s, 4, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'PoE, Power, and Efficiency');
  addTable(s, [
    [hdr('Parameter'), hdr('Cisco C9200CX-12P-2X2G-E'), hdr('Arista CCS-710P-12')],
    ['PoE Standard', 'IEEE 802.3at PoE+', 'IEEE 802.3at PoE+'],
    ['PoE Budget', greenCell('240W'), '234W'],
    ['Max PoE Delta', 'Baseline', redCell('-6W vs Cisco')],
    ['Power Supply', '315W AC internal fixed PSU', amberCell('150W or 280W external AC PSU options')],
    ['Switch Power (excluding PoE)', redCell('Higher system power footprint'), greenCell('46W switching power published')],
    ['Persistent / Fast PoE', greenCell('Perpetual + Fast PoE'), amberCell('Persistent PoE on reboot')],
  ], { y: 1.35, colW: [3.1, 4.5, 4.7] });
  addConclusions(s, [
    'Cisco has a small PoE headroom advantage (240W vs 234W), but the difference is operationally minor in many deployments.',
    'Arista’s published switching power is dramatically lower, making it more power-efficient for quiet edge environments.'
  ], 5.1);
  addFooter(s, 5, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Performance, Throughput, and Scale');
  addTable(s, [
    [hdr('Parameter'), hdr('Cisco C9200CX-12P-2X2G-E'), hdr('Arista CCS-710P-12')],
    ['Switching Capacity', greenCell('68 Gbps'), '64 Gbps'],
    ['Packets / Second', '50.59 Mpps', greenCell('95 Mpps')],
    ['System Memory', '4 GB DRAM', '4 GB system memory'],
    ['Flash Storage', '8 GB flash', '8 GB flash'],
    ['Packet Buffer', greenCell('6 MB'), '2 MB'],
    ['MAC Table', greenCell('32,000'), amberCell('Not clearly published in fetched Arista source set')],
    ['VLAN IDs', '4094', amberCell('Not explicitly confirmed in fetched 710P web page')],
  ], { y: 1.35, colW: [3.0, 4.6, 4.7] });
  addConclusions(s, [
    'Cisco retains an advantage in packet buffer and documented L2 scale visibility.',
    'Arista publishes much higher packets-per-second throughput, suggesting a very efficient compact forwarding path.'
  ], 5.3);
  addFooter(s, 6, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Software, Management, and Telemetry');
  addTable(s, [
    [hdr('Feature'), hdr('Cisco C9200CX-12P-2X2G-E'), hdr('Arista CCS-710P-12')],
    ['Operating System', 'Cisco IOS XE', 'Arista EOS'],
    ['Central Management', 'Catalyst Center / Meraki monitoring', 'CloudVision'],
    ['Streaming Telemetry', greenCell('Model-driven telemetry'), greenCell('Real-time telemetry with CloudVision')],
    ['Flow Visibility', greenCell('Flexible NetFlow'), amberCell('sFlow')],
    ['Programmability', greenCell('NETCONF / RESTCONF / YANG'), greenCell('Extensive EOS APIs / programmability')],
    ['In-Service Updates', amberCell('Cold patching with reboot'), greenCell('In-service maintenance / upgrades highlighted')],
  ], { y: 1.35, colW: [3.0, 4.6, 4.7] });
  addConclusions(s, [
    'Cisco has the richer published flow-monitoring story with Flexible NetFlow.',
    'Arista differentiates with EOS operational model consistency and CloudVision-centric telemetry workflows.'
  ], 5.1);
  addFooter(s, 7, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Security and Campus Feature Positioning');
  addTable(s, [
    [hdr('Feature'), hdr('Cisco C9200CX-12P-2X2G-E'), hdr('Arista CCS-710P-12')],
    ['MACsec', greenCell('AES-256 on C9200CX'), amberCell('Not directly confirmed in fetched 710P source set')],
    ['802.1X', 'Yes', amberCell('Campus security supported, but exact feature list not fully substantiated here')],
    ['Trust Anchor / Secure Boot', greenCell('Cisco Trust Anchor / Secure Boot'), amberCell('Arista security posture present, exact equivalent not fully confirmed here')],
    ['Campus Segmentation', greenCell('SD-Access support'), amberCell('VXLAN/EVPN segmentation highlighted')],
    ['App Visibility', greenCell('NBAR2 / AVC in Cisco family'), redCell('Not positioned as NBAR equivalent')],
  ], { y: 1.45, colW: [3.1, 4.6, 4.6] });
  addConclusions(s, [
    'Cisco’s campus-security and segmentation story is more explicitly documented in the reviewed source set.',
    'Arista is competitive on automation and segmentation architecture, but not all security details were equally explicit on the public pages reviewed.'
  ], 5.0);
  addFooter(s, 8, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'L2 / L3 Protocol and Control-Plane Positioning');
  addTable(s, [
    [hdr('Capability'), hdr('Cisco C9200CX-12P-2X2G-E'), hdr('Arista CCS-710P-12')],
    ['Layer 2 Campus Switching', greenCell('Yes'), greenCell('Yes')],
    ['Static Routing', greenCell('Yes'), amberCell('Expected, but not fully enumerated in fetched web page')],
    ['OSPF / EIGRP / IS-IS / RIP', greenCell('Explicitly listed in Cisco 9200 datasheet family'), amberCell('Not fully enumerated in fetched 710P page')],
    ['BGP / EVPN Positioning', greenCell('Basic BGP on C9200CX from IOS-XE 17.13.1'), greenCell('EVPN segmentation highlighted in 710P positioning')],
    ['PTP', greenCell('IEEE 1588v2 on C9200CX'), amberCell('Not directly substantiated in fetched 710P page')],
  ], { y: 1.45, colW: [3.3, 4.5, 4.5] });
  addConclusions(s, [
    'Cisco publishes a more explicit compact-switch routing and timing feature list in the reviewed documentation.',
    'Arista’s public 710P positioning emphasizes campus segmentation, telemetry, and operational model over detailed per-protocol tables.'
  ], 4.9);
  addFooter(s, 9, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Deployment Fit: Branch, Quiet Spaces, and Edge Closets');
  addTable(s, [
    [hdr('Scenario'), hdr('Best Fit'), hdr('Why')],
    ['Quiet rooms / conference / retail', cell('Tie', { bold: true }), 'Both are compact fanless switches designed for quiet deployment.'],
    ['Need extra copper uplinks', greenCell('Cisco'), 'C9200CX adds 2x 1G copper uplinks.'],
    ['Need lowest switching power draw', greenCell('Arista'), 'Arista publishes lower switching power consumption.'],
    ['Need Cisco campus fabric alignment', greenCell('Cisco'), 'IOS XE, Catalyst Center, Trust Anchor, SD-Access alignment.'],
    ['Need EOS / CloudVision operational model', greenCell('Arista'), 'Single EOS operational experience and Arista campus telemetry approach.'],
  ], { y: 1.45, colW: [3.2, 2.2, 6.9] });
  addConclusions(s, [
    'This is a real head-to-head compact campus comparison: deployment fit is strong on both sides.',
    'The choice is less about hardware shape and more about operational ecosystem: Cisco campus stack vs Arista EOS/CloudVision.'
  ], 4.9);
  addFooter(s, 10, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Critical Caveats and Asymmetries');
  addTable(s, [
    [hdr('Topic'), hdr('Finding')],
    ['Uplink asymmetry', 'Cisco includes 2x 1G copper uplinks in addition to 2x 10G SFP+; Arista 710P-12 does not.'],
    ['PoE budget asymmetry', 'Arista is 6W below Cisco (234W vs 240W) — usually minor, but verify high-density PoE loads.'],
    ['Documentation asymmetry', 'Cisco datasheet exposes detailed L2/L3 scale tables; Arista 710P public pages emphasize operational and platform positioning more than full protocol tables.'],
    ['Security detail asymmetry', 'Cisco MACsec-256 and Trust Anchor are explicit in reviewed sources; equivalent Arista security specifics were not equally explicit in the fetched 710P pages.'],
  ], { y: 1.45, colW: [3.3, 9.0] });
  addConclusions(s, [
    'The deck stays conservative where Arista public documentation was less explicit than Cisco’s datasheet tables.',
    'No unsupported feature was credited to Arista unless it appeared directly in the reviewed official material.'
  ], 4.9);
  addFooter(s, 11, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Recommendation Matrix');
  addTable(s, [
    [hdr('Priority'), hdr('Recommended Platform'), hdr('Why')],
    ['Closest 1:1 compact hardware match', greenCell('Arista CCS-710P-12'), '12x 1G PoE+, 2x 10G SFP+, compact fanless 1RU.'],
    ['Maximum PoE headroom in this comparison', greenCell('Cisco C9200CX-12P-2X2G-E'), '240W vs 234W and richer uplink mix.'],
    ['Best Cisco campus integration', greenCell('Cisco C9200CX-12P-2X2G-E'), 'Catalyst Center, SD-Access, Cisco Trust Anchor, IOS XE feature transparency.'],
    ['Best EOS / Arista operational alignment', greenCell('Arista CCS-710P-12'), 'CloudVision, EOS consistency, efficient compact edge design.'],
    ['Lowest switching power draw', greenCell('Arista CCS-710P-12'), 'Published switching power is substantially lower.'],
  ], { y: 1.45, colW: [3.5, 3.5, 5.3] });
  addConclusions(s, [
    'If you want the cleanest Arista substitute for this exact Cisco compact switch, CCS-710P-12 is the right answer.',
    'Choose Cisco for tighter campus feature integration; choose Arista for EOS/CloudVision alignment and a cleaner low-power compact edge platform.'
  ], 4.9);
  addFooter(s, 12, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Sources & Methodology');
  const sources = [
    ['Cisco Catalyst 9200 Series Data Sheet', 'cisco.com/c/en/us/products/collateral/switches/catalyst-9200-series-switches/nb-06-cat9200-ser-data-sheet-cte-en.html'],
    ['Arista 710P Series Product Page', 'arista.com/en/products/710p-series'],
    ['Arista 710P Datasheet (PDF)', 'arista.com/assets/data/pdf/Datasheets/CCS-710P-Datasheet.pdf'],
    ['Arista 710XP Series Product Page', 'arista.com/en/products/710xp-series'],
    ['Arista 710XP Datasheet', 'arista.com/assets/data/pdf/Datasheets/CCS-710XP-Datasheet.pdf'],
  ];
  let y = 1.45;
  sources.forEach(([name, url]) => {
    s.addText(name, { x: 0.6, y, w: 4.2, h: 0.35, fontFace: FONT, fontSize: 14, bold: true, color: C.navy });
    s.addText(url, { x: 4.8, y, w: 7.9, h: 0.35, fontFace: FONT, fontSize: 12, color: C.blue });
    y += 0.44;
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 4.9, w: 12.3, h: 1.35, rectRadius: 0.04, fill: { color: C.softBg }, line: { color: C.line, pt: 1 } });
  s.addText('Methodology', { x: 0.7, y: 5.1, w: 3.0, h: 0.3, fontFace: FONT, fontSize: 14, bold: true, color: C.navy });
  s.addText('This deck follows the same validated format as the prior Cisco comparison presentations. Arista CCS-710P-12 was selected as the closest official compact PoE near-match to C9200CX-12P-2X2G-E. Where Cisco documentation exposed more detailed scale data than Arista public pages, the deck marks Arista values conservatively rather than guessing.', {
    x: 0.7, y: 5.45, w: 11.8, h: 0.65, fontFace: FONT, fontSize: 13, color: C.sub
  });
  addFooter(s, 13, TOTAL);
}

pptx.writeFile({ fileName: 'C9200CX_vs_Arista_710P_Validated.pptx' });
