const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Carnival Network Engineering';
pptx.company = 'Carnival Corporation';
pptx.subject = 'Cisco C9200CX vs IE3500H Analysis - Corrected & Validated';
pptx.title = 'Cisco C9200CX-12P-2X2G-E vs IE-3500H-12P2MU2XE';
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
const TOTAL = 14;

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
  slide.addText(`Carnival | C9200CX vs IE3500H | Slide ${pg} of ${total}`, {
    x: 0.6, y: 7.05, w: 8, h: 0.25, fontFace: FONT, fontSize: 8, color: C.mut
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
  s.addText('Cisco C9200CX-12P-2X2G-E\nvs. Cisco IE-3500H-12P2MU2XE', {
    x: 0.8, y: 0.8, w: 11.8, h: 2.0, fontFace: FONT, fontSize: 32, bold: true, color: C.hdrTx, lineSpacingMultiple: 1.2
  });
  s.addShape(pptx.ShapeType.line, { x: 0.8, y: 3.0, w: 5.2, h: 0, line: { color: C.blue, pt: 3 } });
  s.addText('Enterprise Compact Access vs. Industrial Heavy-Duty Ethernet', {
    x: 0.8, y: 3.3, w: 11.7, h: 0.6, fontFace: FONT, fontSize: 22, color: C.mut
  });
  s.addText('Context: Same-format validated comparison for harsh-environment deployment decisions\nPrepared For: Carnival', {
    x: 0.8, y: 4.2, w: 9.0, h: 0.9, fontFace: FONT, fontSize: 14, color: C.mut, lineSpacingMultiple: 1.4
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.6, w: 6.5, h: 1.0, rectRadius: 0.06, fill: { color: C.dBlue }, line: { color: C.blue, pt: 1 } });
  s.addText('Validated against official Cisco datasheets and installation guides\nMarch 2026', {
    x: 1.0, y: 5.8, w: 6.1, h: 0.6, fontFace: FONT, fontSize: 13, color: C.hdrTx, bold: true, align: 'center'
  });
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Executive Summary', 'IE-3500H-12P2MU2XE is an industrial near-match, not a 1:1 enterprise peer');
  addTable(s, [
    [hdr('Metric'), hdr('C9200CX-12P-2X2G-E'), hdr('IE-3500H-12P2MU2XE')],
    ['Design Intent', 'Enterprise compact access', greenCell('Industrial / heavy-duty Ethernet')],
    ['Ports', '12x 1G PoE+ + 2x 10G SFP+', amberCell('12x M12 1G PoE + 2x 10G/1G SFP + 2x mGig/4PPoE')],
    ['Power Input', redCell('100-240V AC internal'), greenCell('12-54V DC industrial')],
    ['Ingress Protection', redCell('No IP rating'), greenCell('IP66/IP67')],
    ['Operating Temp', redCell('-5°C to 45°C'), greenCell('-40°C to 75°C')],
    ['PoE Budget', '240W', '240W'],
    ['MACsec', greenCell('AES-256'), amberCell('MACsec-128 on -E license')],
    ['Stacking', redCell('Not supported'), redCell('Not supported')],
    ['Weight', greenCell('6.6 lb (2.99 kg)'), redCell('11.40 lb (5.17 kg)')],
  ], { y: 1.35, colW: [2.9, 4.5, 4.9] });
  addConclusions(s, [
    'C9200CX is lighter and campus-oriented; IE3500H is heavier but purpose-built for harsh environments.',
    'If the real requirement is industrial sealing, DC power, M12, and vibration hardening, IE3500H is the correct class of product.'
  ], 5.8);
  addFooter(s, 2, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Fit Assessment: Is IE-3500H-12P2MU2XE the Closest Match?', 'Validated interpretation of the user-provided candidate');
  addTable(s, [
    [hdr('Question'), hdr('Answer'), hdr('Impact')],
    ['Is it a 12-port PoE Cisco alternative?', greenCell('Yes, partially'), 'It matches the 12-port PoE requirement but adds extra industrial ports.'],
    ['Is it a direct port-for-port peer?', redCell('No'), 'IE3500H has 2 extra mGig/4PPoE ports and M12 industrial connectors.'],
    ['Is it closer in environment / ruggedness?', greenCell('Yes'), 'It is the correct Cisco family if harsh-environment deployment is the real requirement.'],
    ['Is it closer in enterprise access behavior?', amberCell('No'), 'The C9200CX remains the cleaner fit for standard campus/branch access switching.'],
  ], { y: 1.45, colW: [3.7, 2.3, 6.0] });
  addConclusions(s, [
    'This comparison is valid for decision-making, but it compares an enterprise compact switch to an industrial heavy-duty switch.',
    'The deck explicitly treats IE3500H as a harsh-environment near-match, not a strict 1:1 replacement.'
  ], 4.9);
  addFooter(s, 3, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Ports, Uplinks, and Physical Interfaces');
  addTable(s, [
    [hdr('Parameter'), hdr('C9200CX-12P-2X2G-E'), hdr('IE-3500H-12P2MU2XE')],
    ['Primary PoE Ports', '12x 10/100/1000 RJ-45 PoE+', '12x 10/100/1000 M12 X-code PoE'],
    ['Uplink Ports', '2x 10G SFP+', '2x 10G/1G SFP'],
    ['Additional Ports', 'None', amberCell('2x mGig / 4PPoE industrial ports')],
    ['Management / Console', 'RJ45 + USB console / USB-A', 'M12 A-coded 5-pin console'],
    ['Connector Standard', redCell('Enterprise RJ-45'), greenCell('Industrial M12')],
    ['Total Port Count', '14 logical external interfaces', amberCell('16 total interfaces')],
  ], { y: 1.35, colW: [3.2, 4.4, 4.7] });
  addConclusions(s, [
    'IE3500H is not just a rugged C9200CX: it changes the physical interface model with M12 and extra industrial ports.',
    'If field cabling already uses M12 or requires sealed connectors, IE3500H has a major deployment advantage.'
  ], 5.4);
  addFooter(s, 4, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'PoE, Power Architecture, and 4PPoE');
  addTable(s, [
    [hdr('Parameter'), hdr('C9200CX-12P-2X2G-E'), hdr('IE-3500H-12P2MU2XE')],
    ['PoE Standard', 'IEEE 802.3at PoE+ (30W/port)', 'IEEE 802.3at PoE+ + industrial 4PPoE-capable ports'],
    ['PoE Budget', '240W', '240W'],
    ['Max Port Power', '30W', amberCell('Up to 90W on 4PPoE-capable ports')],
    ['Input Power', redCell('100-240V AC'), greenCell('12-54V DC')],
    ['Max System Draw', '315W AC internal PSU', '45W base system without PoE (Cisco published)'],
    ['Deployment Bias', 'Enterprise wiring closets', greenCell('Industrial DC / OT power environments')],
  ], { y: 1.35, colW: [3.1, 4.5, 4.7] });
  addConclusions(s, [
    'Both switches offer 240W PoE budget, but IE3500H aligns better with DC-powered industrial infrastructure.',
    'If the use case needs 4PPoE / high-power industrial endpoints, IE3500H expands beyond standard PoE+.'
  ], 5.2);
  addFooter(s, 5, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Performance, Scale, and Memory');
  addTable(s, [
    [hdr('Parameter'), hdr('C9200CX-12P-2X2G-E'), hdr('IE-3500H-12P2MU2XE')],
    ['Switching Capacity', greenCell('68 Gbps'), amberCell('Not officially published for this SKU')],
    ['Forwarding Rate', '50.59 Mpps', greenCell('Line rate for all ports / packet sizes')],
    ['MAC Address Table', greenCell('32,000'), '24,000'],
    ['VLAN IDs', '4,094 / 4,096 class', '4,096'],
    ['DRAM', amberCell('Not officially published per reviewed source set'), '8 GB'],
    ['Flash / User Storage', amberCell('Not officially published per reviewed source set'), '5.1 GB user accessible'],
    ['IPv4 Routes', 'Approx. 3,000 class (license/template dependent)', greenCell('7,000 IPv4')],
  ], { y: 1.35, colW: [3.0, 4.6, 4.7] });
  addConclusions(s, [
    'C9200CX leads in L2 MAC scale; IE3500H leads in published industrial routing scale and memory footprint.',
    'Cisco does not publish every compact C9200CX memory detail cleanly for this exact SKU, so the deck marks those as not officially substantiated.'
  ], 5.4);
  addFooter(s, 6, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Security and Encryption');
  addTable(s, [
    [hdr('Feature'), hdr('C9200CX-12P-2X2G-E'), hdr('IE-3500H-12P2MU2XE')],
    ['MACsec', greenCell('AES-256 on C9200CX'), amberCell('MACsec-128 on Network Essentials (-E)')],
    ['Secure Boot / Trust Anchor', greenCell('Yes'), greenCell('Yes')],
    ['802.1X', 'Yes', 'Yes'],
    ['ACLs', 'Yes', 'Yes'],
    ['Campus Security Features', greenCell('Stronger campus/security fabric orientation'), amberCell('Industrial security focus rather than campus fabric')],
    ['SD-Access / DNA Center', greenCell('Supported'), greenCell('Supported with DNA Center')],
  ], { y: 1.35, colW: [3.2, 4.4, 4.7] });
  addConclusions(s, [
    'C9200CX has the cleaner encryption advantage because AES-256 MACsec is standard for the reviewed compact model.',
    'The IE model in the user table ends with -E, so MACsec-256 should not be claimed unless the license changes to Network Advantage.'
  ], 5.1);
  addFooter(s, 7, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'L2 / L3 Protocols and Timing');
  addTable(s, [
    [hdr('Feature'), hdr('C9200CX-12P-2X2G-E'), hdr('IE-3500H-12P2MU2XE')],
    ['Layer 2', 'VLANs, STP, RSTP, MST, LACP', 'VLANs, STP, RSTP, MST, LACP'],
    ['Static Routing', 'Yes', 'Yes'],
    ['OSPF', 'Yes (license dependent)', 'Yes'],
    ['EIGRP', 'Yes (license dependent)', 'Yes'],
    ['BGP', amberCell('Limited / platform-license dependent'), greenCell('Supported in reviewed IE3500H source set')],
    ['PTP (IEEE 1588v2)', greenCell('Yes'), greenCell('Yes')],
    ['Industrial Protocols', redCell('No'), greenCell('EtherNet/IP, PROFINET, MODBUS / SCADA class support')],
    ['Redundancy Protocol Bias', 'Campus bridging/routing', greenCell('Industrial ring / OT redundancy focus')],
  ], { y: 1.35, colW: [3.0, 4.5, 4.8] });
  addConclusions(s, [
    'Both platforms support mainstream L2/L3 and PTP, but IE3500H adds industrial protocol relevance that the campus Catalyst line does not target.',
    'For OT synchronization and industrial interoperability, IE3500H is the more natural fit.'
  ], 5.6);
  addFooter(s, 8, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Environmental Hardening and Industrial Suitability');
  addTable(s, [
    [hdr('Criterion'), hdr('C9200CX-12P-2X2G-E'), hdr('IE-3500H-12P2MU2XE')],
    ['Ingress Protection', redCell('No IP rating'), greenCell('IP66 / IP67')],
    ['Operating Temperature', redCell('-5°C to 45°C'), greenCell('-40°C to 75°C')],
    ['Shock / Vibration', redCell('Not industrial rated'), greenCell('IEC 60068 industrial shock / vibration tested')],
    ['Salt Mist / Corrosion', amberCell('Not positioned for marine/industrial exposure'), amberCell('Salt mist testing cited as in progress in reviewed source')],
    ['Mounting Style', 'Enterprise compact / cabinet', greenCell('Industrial wall mount / field deployment')],
    ['Power Environment', redCell('AC-only'), greenCell('DC-ready industrial')],
  ], { y: 1.35, colW: [3.0, 4.5, 4.8] });
  addConclusions(s, [
    'IE3500H is the only model in this comparison truly designed for harsh-environment deployment.',
    'If the deployment is marine, outdoor, rail, utility, or factory-like, C9200CX is the wrong product class.'
  ], 5.1);
  addFooter(s, 9, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Industrial Certifications and Deployment Readiness');
  addTable(s, [
    [hdr('Certification / Readiness Item'), hdr('C9200CX-12P-2X2G-E'), hdr('IE-3500H-12P2MU2XE')],
    ['IEC / UL Safety', 'Enterprise safety certifications', greenCell('Industrial safety certifications')],
    ['IEEE 1613 / Utility Class', redCell('No'), amberCell('Family support / in-progress references depending document')],
    ['EN50155 / Rail', redCell('No'), greenCell('Yes')],
    ['E-Mark / Automotive', redCell('No'), greenCell('Yes')],
    ['ATEX / IECEx / Harsh OT Positioning', redCell('No industrial positioning'), greenCell('Industrial / heavy-duty positioning')],
    ['Marine Deployment Readiness', redCell('Not suitable without major enclosure adaptation'), greenCell('Much closer fit for marine/harsh field scenarios')],
  ], { y: 1.35, colW: [3.5, 4.0, 4.8] });
  addConclusions(s, [
    'IE3500H carries the regulatory and environmental posture expected from an industrial Ethernet switch.',
    'C9200CX remains an enterprise compact switch even when its port count looks superficially similar.'
  ], 5.1);
  addFooter(s, 10, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Weight, Dimensions, and Physical Trade-Offs');
  addTable(s, [
    [hdr('Parameter'), hdr('C9200CX-12P-2X2G-E'), hdr('IE-3500H-12P2MU2XE')],
    ['Weight', greenCell('6.6 lb (2.99 kg)'), redCell('11.40 lb (5.17 kg)')],
    ['Dimensions', '4.4 x 26.9 x 24.4 cm', '26.57 x 27.69 x 8.53 cm'],
    ['Noise', greenCell('Fanless / silent'), greenCell('Fanless')],
    ['Cabinet Impact', greenCell('Lower weight, easier retrofits'), amberCell('Heavier but more rugged package')],
    ['Connector Robustness', redCell('RJ-45 vulnerable in harsh field use'), greenCell('M12 improves field durability')],
  ], { y: 1.5, colW: [3.2, 4.3, 4.8] });
  addConclusions(s, [
    'C9200CX clearly wins on weight and conventional cabinet integration.',
    'IE3500H accepts a weight penalty in exchange for sealed construction and field-grade connector durability.'
  ], 4.7);
  addFooter(s, 11, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Critical Differences from the User-Provided Table', 'Corrections applied before generating this deck');
  addTable(s, [
    [hdr('User Table Claim'), hdr('Verdict'), hdr('Correction')],
    ['C9200CX switching capacity = 128 Gbps', redCell('False / contradicted'), 'Use 68 Gbps for this compact SKU; 128 Gbps is not the correct figure for this model.'],
    ['IE-3500H-12P2MU2XE = 12 ports total', redCell('False'), 'It includes 12 main M12 PoE ports plus 2 SFP uplinks and 2 additional mGig/4PPoE interfaces.'],
    ['IE -E license has AES-256 MACsec', amberCell('Misleading'), 'For the -E (Network Essentials) variant, reviewed evidence supports MACsec-128, not MACsec-256.'],
    ['C9200CX and IE3500H are direct peers', amberCell('Oversimplified'), 'They overlap in PoE compact switching, but one is enterprise access and the other is heavy-duty industrial Ethernet.'],
  ], { y: 1.45, colW: [3.8, 2.0, 6.5], rowH: 0.46 });
  addConclusions(s, [
    'The biggest factual fixes are the C9200CX switching-capacity number and the IE3500H total-interface count.',
    'The licensing detail matters: MACsec-256 should not be attributed to the IE -E variant without Network Advantage.'
  ], 5.2);
  addFooter(s, 12, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Recommendation Matrix');
  addTable(s, [
    [hdr('Scenario'), hdr('Recommended Model'), hdr('Why')],
    ['Standard enterprise branch / cabin / office access', greenCell('C9200CX-12P-2X2G-E'), 'Lighter, simpler, AC powered, enterprise access feature set.'],
    ['Harsh environment / marine exterior / industrial zone', greenCell('IE-3500H-12P2MU2XE'), 'IP67, M12, DC power, shock/vibration hardening, industrial certifications.'],
    ['Need conventional RJ-45 access and low weight', greenCell('C9200CX-12P-2X2G-E'), 'Best fit for standard IT infrastructure.'],
    ['Need M12 connectors and sealed enclosure', greenCell('IE-3500H-12P2MU2XE'), 'This is the core reason to choose the industrial platform.'],
    ['Need strongest MACsec posture on the compared SKUs', greenCell('C9200CX-12P-2X2G-E'), 'Reviewed configuration has native AES-256 vs IE -E MACsec-128.'],
    ['Need OT / industrial protocol alignment', greenCell('IE-3500H-12P2MU2XE'), 'Industrial Ethernet protocols and field deployment posture make it the right class of product.'],
  ], { y: 1.35, colW: [3.6, 3.3, 5.4], rowH: 0.44 });
  addConclusions(s, [
    'Choose C9200CX when the environment is enterprise IT; choose IE3500H when the environment itself is the primary constraint.',
    'This is less about raw port counts and more about deployment physics: AC vs DC, RJ-45 vs M12, office vs harsh field.'
  ], 5.2);
  addFooter(s, 13, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Sources & Methodology');
  const sources = [
    ['Cisco Catalyst 9200 Series Data Sheet', 'cisco.com/c/en/us/products/collateral/switches/catalyst-9200-series-switches/nb-06-cat9200-ser-data-sheet-cte-en.html'],
    ['Cisco IE3500 Heavy Duty Series Data Sheet', 'cisco.com/c/en/us/products/collateral/networking/industrial-switches/ie3500-heavy-duty-series/ie3500-heavy-duty-series-ds.html'],
    ['Cisco IE3500 Rugged Series Data Sheet', 'cisco.com/c/en/us/products/collateral/networking/industrial-switches/ie3500-rugged-series/ie3500-rugged-series-ds.html'],
    ['Cisco IE3500 Hardware Installation Guide', 'cisco.com/c/en/us/td/docs/IIOT/switches/ie35xx/hw-install-guide/b-ie35xx-hig'],
    ['Cisco C9200CX Hardware Install Guide', 'cisco.com/c/en/us/td/docs/switches/lan/catalyst9200/hardware/install/b-c9200cx-hig'],
  ];
  let y = 1.45;
  sources.forEach(([name, url]) => {
    s.addText(name, { x: 0.6, y, w: 4.2, h: 0.35, fontFace: FONT, fontSize: 14, bold: true, color: C.navy });
    s.addText(url, { x: 4.8, y, w: 7.9, h: 0.35, fontFace: FONT, fontSize: 12, color: C.blue });
    y += 0.44;
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 4.9, w: 12.3, h: 1.35, rectRadius: 0.04, fill: { color: C.softBg }, line: { color: C.line, pt: 1 } });
  s.addText('Methodology', { x: 0.7, y: 5.1, w: 3, h: 0.3, fontFace: FONT, fontSize: 14, bold: true, color: C.navy });
  s.addText('This deck uses the same validated format as the prior C9200CX replacement presentation. The user-provided comparison table was treated as input only, then corrected against official Cisco datasheets and installation guides. Unsupported values were either removed or labeled as not officially substantiated. IE-3500H-12P2MU2XE is treated as an industrial near-match rather than a strict 1:1 peer.', {
    x: 0.7, y: 5.45, w: 11.8, h: 0.65, fontFace: FONT, fontSize: 13, color: C.sub
  });
  addFooter(s, 14, TOTAL);
}

pptx.writeFile({ fileName: 'C9200CX_vs_IE3500H_Validated.pptx' });
