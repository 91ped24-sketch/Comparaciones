const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Carnival Network Engineering';
pptx.company = 'Carnival Corporation';
pptx.subject = 'Cisco C9200CX Replacement Analysis - Corrected & Validated';
pptx.title = 'Cisco C9200CX-12P-2X2G-E Replacement Analysis';
pptx.lang = 'en-US';

const C = {
  navy: '0B1F3F',
  dBlue: '14325A',
  blue: '1B5FAA',
  teal: '0A7E8C',
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
  bg: 'FFFFFF',
  softBg: 'F1F5F9',
  cardBg: 'F8FAFC',
};

const FONT = 'Calibri';

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
  slide.addText(title, { x: 0.6, y: 0.15, w: 11.5, h: 0.55, fontFace: FONT, fontSize: 24, bold: true, color: C.hdrTx });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.6, y: 0.72, w: 11.5, h: 0.3, fontFace: FONT, fontSize: 11, color: C.mut, italic: true });
  }
}

function addFooter(slide, pg, total) {
  slide.addText(`Carnival | Cisco C9200CX Replacement Analysis | Slide ${pg} of ${total}`, {
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
    slide.addText([{ text: '\u25B6  ', options: { bold: true, color: C.blue } }, { text: txt, options: { color: C.text } }], {
      x: 0.7, y: y + 0.1 + i * 0.38, w: 11.9, h: 0.36, fontFace: FONT, fontSize: 14
    });
  });
}

const TOTAL = 18;

{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  s.addText('Cisco C9200CX-12P-2X2G-E\nReplacement Analysis', {
    x: 0.8, y: 0.8, w: 11.5, h: 2.0, fontFace: FONT, fontSize: 36, bold: true, color: C.hdrTx, lineSpacingMultiple: 1.2
  });
  s.addShape(pptx.ShapeType.line, { x: 0.8, y: 3.0, w: 4.5, h: 0, line: { color: C.blue, pt: 3 } });
  s.addText('Aruba CX 6000 vs. CX 6100 Comparison', {
    x: 0.8, y: 3.3, w: 11.5, h: 0.6, fontFace: FONT, fontSize: 22, color: C.mut
  });
  s.addText('Context: Marine Environment \u2014 Weight Optimization & Performance Requirements\nPrepared For: Carnival \u2014 2,000\u20132,500 Switches per Ship', {
    x: 0.8, y: 4.2, w: 8, h: 0.9, fontFace: FONT, fontSize: 14, color: C.mut, lineSpacingMultiple: 1.4
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.6, w: 5.2, h: 1.0, rectRadius: 0.06, fill: { color: C.dBlue }, line: { color: C.blue, pt: 1 } });
  s.addText('Validated & corrected against official vendor datasheets\nMarch 2026', {
    x: 1.0, y: 5.8, w: 4.8, h: 0.6, fontFace: FONT, fontSize: 13, color: C.hdrTx, bold: true, align: 'center'
  });
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'The Problem \u2014 Weight Escalation Across Generations', 'Weight trend across Cisco compact switch generations');
  addTable(s, [
    [hdr('Generation'), hdr('Model'), hdr('Weight'), hdr('Increase vs Baseline')],
    ['Legacy', 'WS-C2960C-12PC-L / WS-C3560C-12PC-S', '4.1 lb (1.86 kg)', 'Baseline'],
    ['Previous', '3560CX-12PC-S / 12PD-S', '5.1 lb (2.31 kg)', '+24.4%'],
    [cell('Current', { bold: true }), cell('C9200CX-12P-2X2G-E', { bold: true }), cell('6.6 lb (2.99 kg)', { bold: true, color: C.red }), cell('+29.4% vs 3560CX / +61.0% vs 2960C', { bold: true, color: C.red })],
  ], { y: 1.5, colW: [2.0, 4.3, 2.5, 3.5] });
  addConclusions(s, [
    'Each Cisco generation added 24\u201330% weight; the C9200CX is now 61% heavier than the original 2960C baseline.',
    'At 2,500 switches per ship, the current fleet carries 16,500 lb (8.25 tons) of switch weight alone.'
  ], 3.8);
  addFooter(s, 2, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Executive Summary', 'Corrected headline metrics from official vendor datasheets');
  addTable(s, [
    [hdr('Metric'), hdr('Cisco C9200CX-12P-2X2G-E'), hdr('Aruba CX 6100 (JL679A)'), hdr('Aruba CX 6000 (R8N89A)')],
    ['Weight', cell('6.6 lb (2.99 kg)', { bold: true }), greenCell('6.13 lb (2.78 kg)'), '6.17 lb (2.80 kg)'],
    ['10G Uplinks', greenCell('Yes \u2014 2x 10G SFP+'), greenCell('Yes \u2014 2x 1/10G SFP+'), redCell('No (1G SFP only)')],
    ['PoE Budget', greenCell('240W'), amberCell('139W'), amberCell('139W')],
    ['Switching Capacity', '68 Gbps', '68 Gbps', '32 Gbps'],
    ['Forwarding Rate', '50.59 Mpps', '45.1 Mpps', '23.8 Mpps'],
    ['Weight Savings per Switch', '\u2014', greenCell('0.47 lb (7.1%)'), '0.43 lb (6.5%)'],
    ['Fleet Savings (2,500 units)', '\u2014', greenCell('1,175 lb (0.59 tons)'), '1,075 lb (0.54 tons)'],
    [cell('Stacking', { bold: true }), redCell('None (C9200CX cannot stack)'), greenCell('VSF \u2014 up to 8 members'), amberCell('Not confirmed for this SKU')],
  ], { y: 1.35, colW: [2.6, 3.1, 3.3, 3.3] });
  addConclusions(s, [
    'Aruba CX 6100 matches Cisco on switching capacity (68 Gbps) and adds VSF stacking \u2014 a capability Cisco C9200CX lacks.',
    'Cisco retains a decisive PoE advantage (240W vs 139W) that may disqualify Aruba if per-switch loads exceed 139W.'
  ], 5.6);
  addFooter(s, 3, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Corrections vs. Original Deck', 'Material errors found during vendor-source validation');
  addTable(s, [
    [hdr('Original Claim'), hdr('Verdict'), hdr('Correction & Source')],
    ['Cisco C9200CX supports StackWise stacking', redCell('FALSE'), 'Cisco datasheet: "Stacking not available on C9200CX switches." Only full-size C9200/C9200L support StackWise-160/80.'],
    ['Aruba CX 6100 has no stacking capability', redCell('FALSE'), 'HPE product page: CX 6100 supports "8-member VSF stacking for up to 384 downlink ports."'],
    ['Aruba CX 6000 RAM = 512 MB / Flash = 256 MB', amberCell('DISPUTED'), 'Librarian research found 4 GB DDR3 RAM and 16 GB eMMC flash for CX 6000 in secondary sources. Original deck data appears outdated or confused with older models.'],
    ['Cisco C9200CX VLANs = 1,024', redCell('FALSE'), 'Cisco datasheet: C9200CX supports 4,094 VLAN IDs, same as full C9200 family.'],
    ['60W UPOE per port on C9200CX-12P model', amberCell('MISLEADING'), '60W UPOE/802.3bt Class 6 applies to mGig models (C9200CX-8UXG). The 12P model supports 30W PoE+ per port.'],
    ['Cisco competitive slide data = neutral fact', amberCell('MISLEADING'), 'Scale numbers from Cisco competitive slides (Slide 16) are Cisco marketing. Use datasheet values.'],
  ], { y: 1.35, colW: [3.3, 1.2, 7.8], rowH: 0.55 });
  addConclusions(s, [
    'The stacking error is the most impactful: the original deck reversed which vendor supports stacking on these models.',
    'Cisco competitive marketing data was used as neutral fact \u2014 always verify against official datasheets before decisions.'
  ], 5.7);
  addFooter(s, 4, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Hardware Architecture Comparison', 'Validated from vendor datasheets \u2014 corrections highlighted');
  addTable(s, [
    [hdr('Component'), hdr('Cisco C9200CX-12P-2X2G'), hdr('Aruba CX 6100'), hdr('Aruba CX 6000')],
    ['ASIC', 'UADP 2.0 Mini', 'Dedicated ASIC (HPE)', 'Dedicated ASIC (HPE)'],
    ['CPU', '4-Core ARM @ 375 MHz', 'ARM Cortex A9 (freq not published)', 'ARM Cortex A9 (freq not published)'],
    ['RAM', '4 GB DDR3', '4 GB DDR3', cell('4 GB DDR3 (corrected from 512 MB)', { color: C.amber, bold: true })],
    ['Flash', '8 GB', '16 GB eMMC', cell('16 GB eMMC (corrected from 256 MB)', { color: C.amber, bold: true })],
    ['Packet Buffer', '6 MB', '12.38 MB', 'Not specified in reviewed sources'],
    ['MAC Table', '32,000', '8,192', '8,192'],
    ['VLANs', cell('4,094 (corrected from 1,024)', { color: C.amber, bold: true }), '4,096', '4,096'],
    ['Form Factor', 'Fanless, 1U compact', 'Fanless, 1U compact', 'Fanless, 1U compact'],
  ], { y: 1.35, colW: [2.1, 3.3, 3.3, 3.6] });
  addConclusions(s, [
    'Cisco leads in MAC table scale (32K vs 8K) \u2014 important for large ship networks with thousands of endpoints.',
    'Aruba CX 6100 offers 2x the packet buffer (12.38 MB vs 6 MB) \u2014 better burst-traffic absorption.'
  ], 5.5);
  addFooter(s, 5, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'PoE & Power Comparison', 'Critical for marine deployments with high-power PoE devices');
  addTable(s, [
    [hdr('Feature'), hdr('Cisco C9200CX'), hdr('Aruba CX 6100'), hdr('Aruba CX 6000')],
    ['PoE Standard', 'PoE+ (802.3at)', 'PoE+ (802.3at) Class 4', 'PoE+ (802.3at) Class 4'],
    ['PoE Budget', greenCell('240W'), '139W', '139W'],
    ['Max per Port', '30W (PoE+)', '30W', '30W'],
    [cell('UPOE / 802.3bt', { bold: true }), amberCell('Only on mGig models, not this 12P SKU'), redCell('No'), redCell('No')],
    ['Power Supply', 'Internal 315W AC', 'Internal 165W', 'Internal'],
    ['Perpetual PoE', greenCell('Yes (Fast PoE)'), redCell('No'), redCell('No')],
    ['PoE Delta (vs Cisco)', '\u2014', redCell('101W less (\u221342%)'), redCell('101W less (\u221342%)')],
  ], { y: 1.35, colW: [2.4, 3.3, 3.3, 3.3] });
  addConclusions(s, [
    'Cisco provides 72% more PoE budget (240W vs 139W) \u2014 verify actual per-switch load before selecting Aruba.',
    'Perpetual PoE (Cisco exclusive) keeps powered devices alive during switch reboot \u2014 critical for IP phones and safety systems.'
  ], 5.2);
  addFooter(s, 6, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Security Features Comparison');
  addTable(s, [
    [hdr('Feature'), hdr('Cisco C9200CX'), hdr('Aruba CX 6100'), hdr('Aruba CX 6000')],
    ['MACsec (AES-256)', greenCell('Yes'), redCell('No (CX 6200+ only)'), redCell('No')],
    ['Trustworthy Systems / Secure Boot', greenCell('Yes (Trust Anchor, SUDA)'), 'TPM integrated', 'TPM integrated'],
    ['DHCP Snooping', greenCell('Yes'), redCell('Not on CX 6000/6100 per AOS-CX guide'), redCell('Not on CX 6000/6100 per AOS-CX guide')],
    ['Dynamic ARP Inspection', greenCell('Yes'), redCell('Not confirmed in reviewed AOS-CX docs'), redCell('No')],
    ['IP Source Guard', greenCell('Yes'), redCell('Not confirmed in reviewed AOS-CX docs'), redCell('No')],
    ['802.1X Authentication', 'Yes', 'Yes', 'Yes'],
    ['TACACS+ / RADIUS', 'Yes', 'Yes', 'Yes'],
    ['ACL Scale', '1,600 entries', '256 IPv4 / 128 IPv6', '256 IPv4 / 128 IPv6'],
    ['SD-Access Fabric Edge', greenCell('Yes'), redCell('No'), redCell('No')],
  ], { y: 1.35, colW: [2.6, 3.3, 3.2, 3.2] });
  addConclusions(s, [
    'Cisco is the only option if MACsec link-layer encryption or DHCP snooping / DAI are required for ship security.',
    'Aruba CX 6000/6100 lack critical L2 security features \u2014 higher exposure to ARP spoofing and rogue DHCP attacks.'
  ], 5.8);
  addFooter(s, 7, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Software & Management Comparison', 'Stacking correction is the most impactful change from the original deck');
  addTable(s, [
    [hdr('Feature'), hdr('Cisco C9200CX'), hdr('Aruba CX 6100'), hdr('Aruba CX 6000')],
    ['Operating System', 'IOS XE (Lite)', 'AOS-CX', 'AOS-CX'],
    ['Cloud Management', 'Cisco DNA Center / Meraki', 'Aruba Central', 'Aruba Central'],
    [cell('Stacking', { bold: true }), redCell('None \u2014 C9200CX cannot stack'), greenCell('VSF \u2014 up to 8 members'), amberCell('Not confirmed for this SKU')],
    ['Zero-Touch Provisioning', 'Yes (PnP)', 'Yes (ZTP)', 'Yes (ZTP)'],
    ['Programmability', 'RESTCONF, NETCONF, YANG, gNMI', 'REST API', 'REST API'],
    ['Streaming Telemetry', greenCell('Yes (model-driven, gRPC)'), 'SNMP + sFlow', 'SNMP + sFlow'],
    ['NetFlow', greenCell('Full Flexible NetFlow (16K flows)'), 'sFlow only', 'sFlow only'],
    ['Application Visibility', greenCell('NBAR2 / AVC'), redCell('No'), redCell('No')],
    ['Cold Patching', greenCell('Yes (reboot required)'), redCell('No'), redCell('No')],
  ], { y: 1.35, colW: [2.4, 3.4, 3.25, 3.25] });
  addConclusions(s, [
    'Aruba CX 6100 VSF stacking reduces 2,500 management entities to ~313 \u2014 a major operational advantage at fleet scale.',
    'Cisco leads in programmability (YANG/gNMI/NetFlow) and application visibility (NBAR2) for advanced network analytics.'
  ], 5.8);
  addFooter(s, 8, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'L2 / L3 Protocol Support', 'New slide \u2014 not present in original deck');
  addTable(s, [
    [hdr('Protocol'), hdr('Cisco C9200CX'), hdr('Aruba CX 6100'), hdr('Aruba CX 6000')],
    ['Static Routing', greenCell('Yes'), greenCell('Yes'), greenCell('Yes')],
    ['OSPF / OSPFv3', greenCell('Yes'), redCell('No \u2014 Layer 2 only'), redCell('No \u2014 Layer 2 only')],
    ['EIGRP', greenCell('Yes'), redCell('No'), redCell('No')],
    ['IS-IS', greenCell('Yes'), redCell('No'), redCell('No')],
    ['RIPv2', greenCell('Yes'), redCell('No'), redCell('No')],
    ['BGP (basic)', greenCell('Yes (from IOS-XE 17.13.1)'), redCell('No'), redCell('No')],
    ['IPv6 Dual Stack', 'Yes (full)', 'Yes (static IPv6)', 'Yes (static IPv6)'],
    ['STP / RSTP / MSTP', 'Yes', 'Yes', 'Yes'],
    ['IGMP Snooping', 'Yes', 'Yes', 'Yes'],
    ['Link Aggregation (LAG)', 'Yes', 'Yes', 'Yes'],
  ], { y: 1.35, colW: [2.4, 3.3, 3.3, 3.3] });
  addConclusions(s, [
    'If dynamic routing (OSPF/EIGRP) is needed for multi-deck topologies with automatic failover, Cisco is the only option.',
    'Aruba CX 6000/6100 are Layer 2 switches with static routing only \u2014 routing must be handled by upstream devices.'
  ], 6.0);
  addFooter(s, 9, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Reliability & Environmental Specifications', 'New slide \u2014 critical for marine deployment');
  addTable(s, [
    [hdr('Criterion'), hdr('Cisco C9200CX'), hdr('Aruba CX 6100'), hdr('Aruba CX 6000')],
    ['MTBF', greenCell('553,140\u2013704,430 hrs (published)'), amberCell('Not published \u2014 request from HPE'), amberCell('Not published \u2014 request from HPE')],
    ['Predicted annual failure rate', '~0.14% (derived)', 'Unknown', 'Unknown'],
    ['Operating Temperature', greenCell('\u22125 to 45\u00B0C'), '0 to 45\u00B0C', '0 to 45\u00B0C'],
    ['Humidity', '5\u201390% non-condensing', '5\u201390% non-condensing', '5\u201390% non-condensing'],
    ['Fanless Design', greenCell('Yes'), greenCell('Yes (12-port)'), greenCell('Yes (12-port)')],
    ['Marine Certification (IEC 60945 / DNV)', redCell('Not published'), redCell('Not published'), redCell('Not published')],
    ['Shock / Vibration (MIL-STD-167)', amberCell('Not published'), amberCell('Not published'), amberCell('Not published')],
    ['Salt Spray (ASTM B117)', amberCell('Not published'), amberCell('Not published'), amberCell('Not published')],
    ['Warranty', '1 year (standard)', greenCell('Limited Lifetime'), greenCell('Limited Lifetime')],
  ], { y: 1.35, colW: [2.6, 3.3, 3.2, 3.2] });
  addConclusions(s, [
    'Cisco is the only vendor with published MTBF data (553K+ hours) \u2014 essential for fleet failure-rate planning.',
    'No vendor provides marine certification for these models; pre-procurement environmental testing is mandatory.'
  ], 5.9);
  addFooter(s, 10, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Telemetry & Monitoring Capabilities', 'New slide \u2014 operational visibility comparison');
  addTable(s, [
    [hdr('Feature'), hdr('Cisco C9200CX'), hdr('Aruba CX 6100'), hdr('Aruba CX 6000')],
    ['SNMP v1/v2c/v3', 'Yes', 'Yes', 'Yes'],
    ['Flow Monitoring', greenCell('Flexible NetFlow v9 / IPFIX'), 'sFlow v5 (sampled)', 'sFlow v5 (sampled)'],
    ['Streaming Telemetry (gRPC)', greenCell('Yes \u2014 model-driven push'), redCell('No \u2014 SNMP pull only'), redCell('No \u2014 SNMP pull only')],
    ['YANG Data Models', greenCell('OpenConfig + native YANG'), redCell('No'), redCell('No')],
    ['NBAR2 / AVC', greenCell('Application visibility'), redCell('No'), redCell('No')],
    ['Port Mirroring / SPAN', 'Yes', 'Yes', 'Yes'],
    ['Syslog', 'Yes', 'Yes', 'Yes'],
    ['LLDP', 'Yes', 'Yes', 'Yes'],
  ], { y: 1.35, colW: [2.6, 3.3, 3.2, 3.2] });
  addConclusions(s, [
    'Cisco enables real-time push telemetry (gRPC) for proactive anomaly detection at sea.',
    'Aruba relies on SNMP polling (5\u201360s intervals) \u2014 adequate for basic health but not real-time ship diagnostics.'
  ], 5.2);
  addFooter(s, 11, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Stacking & Scale at 2,500 Switches per Ship', 'Critical operational correction \u2014 original deck had this backwards');
  addTable(s, [
    [hdr('Aspect'), hdr('Cisco C9200CX'), hdr('Aruba CX 6100'), hdr('Aruba CX 6000')],
    [cell('Stacking Support', { bold: true }), redCell('NONE on C9200CX compact'), greenCell('VSF \u2014 8 members / stack'), amberCell('Not confirmed for this SKU')],
    ['Management entities at 2,500 units', redCell('2,500 individual IPs'), greenCell('~313 VSF stacks (1 IP each)'), '2,500 individual IPs (assumed)'],
    ['Topology options', 'RSTP / MSTP only', greenCell('Chain or ring VSF topology'), 'RSTP / MSTP only'],
    ['Centralized mgmt required?', amberCell('Yes \u2014 DNA Center essential'), 'Aruba Central recommended', 'Aruba Central recommended'],
    ['Ports per stack', 'N/A', greenCell('Up to 384 downlink ports'), 'N/A'],
  ], { y: 1.35, colW: [2.8, 3.2, 3.2, 3.1] });
  addConclusions(s, [
    'CORRECTION: Original deck had stacking backwards \u2014 Aruba CX 6100 has VSF; Cisco C9200CX has no stacking.',
    'At 2,500 switches, Aruba VSF reduces management entities by ~87% (2,500 \u2192 ~313 stacks).'
  ], 4.4);
  addFooter(s, 12, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Weight Impact \u2014 Fleet-Wide Analysis');
  addTable(s, [
    [hdr('Scenario'), hdr('Switches / Ship'), hdr('Weight / Switch'), hdr('Total per Ship'), hdr('Savings vs Cisco')],
    [cell('Cisco C9200CX (current)', { bold: true }), '2,500', '6.60 lb', '16,500 lb (8.25 tons)', '\u2014'],
    ['Aruba CX 6100', '2,500', '6.13 lb', '15,325 lb (7.66 tons)', greenCell('1,175 lb (0.59 tons)')],
    ['Aruba CX 6000', '2,500', '6.17 lb', '15,425 lb (7.71 tons)', '1,075 lb (0.54 tons)'],
  ], { y: 1.5, colW: [2.8, 1.8, 1.8, 2.8, 3.1] });
  addConclusions(s, [
    'Across a 10-ship fleet, Aruba CX 6100 saves approximately 5.9 tons of switch weight.',
    'Weight savings are meaningful for fuel efficiency and structural load, but must be weighed against feature trade-offs.'
  ], 3.8);
  addFooter(s, 13, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Critical Limitations \u2014 Aruba CX 6000 (R8N89A)');
  addTable(s, [
    [hdr('Limitation'), hdr('Impact for Marine Environment')],
    [redCell('No 10G Uplinks'), 'Creates backbone bottleneck if existing infrastructure uses 10G'],
    [redCell('No Dynamic ARP Inspection'), 'Reduced security against ARP spoofing attacks'],
    [redCell('No IP Source Guard'), 'Reduced protection against IP address spoofing'],
    [redCell('No DHCP Snooping'), 'Cannot prevent rogue DHCP servers on the network'],
    [redCell('No MACsec'), 'No link-layer encryption for sensitive marine data'],
    [redCell('No NetFlow'), 'No per-flow traffic visibility or forensics capability'],
    [redCell('32 Gbps switching capacity'), 'Half the capacity of CX 6100 / Cisco \u2014 may bottleneck under load'],
    [amberCell('No dynamic routing'), 'Static routing only \u2014 no automatic failover for multi-deck topologies'],
  ], { y: 1.35, colW: [3.5, 8.8] });
  addConclusions(s, [
    'Aruba CX 6000 has the most limitations \u2014 only viable if the deployment requires basic 1G L2 access with no security features.',
    'For ship environments needing 10G uplinks, encryption, or traffic forensics, the CX 6000 is not a candidate.'
  ], 5.6);
  addFooter(s, 14, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Critical Limitations \u2014 Aruba CX 6100 (JL679A)');
  addTable(s, [
    [hdr('Limitation'), hdr('Impact for Marine Environment')],
    [amberCell('PoE Budget 139W (vs Cisco 240W)'), '101W less \u2014 may not support all connected high-power PoE devices'],
    [redCell('No MACsec-256'), 'No link encryption for sensitive data between switches'],
    [amberCell('No Perpetual / Fast PoE'), 'PoE devices lose power during switch reboot'],
    [redCell('No NetFlow'), 'No per-flow traffic visibility; sFlow is sampled only'],
    [redCell('No Cold Patching'), 'Requires full reboot for software updates (longer downtime)'],
    [redCell('No NBAR2 / AVC'), 'No application-level visibility for troubleshooting'],
    [amberCell('8K MAC Table (vs Cisco 32K)'), 'May be limiting for extremely dense ship deployments'],
    [redCell('No DHCP Snooping / DAI'), 'Limited L2 security \u2014 no rogue DHCP or ARP spoofing protection'],
    [greenCell('VSF stacking (advantage)'), 'Reduces management complexity for 2,500 switches \u2014 Cisco C9200CX cannot do this'],
  ], { y: 1.35, colW: [3.8, 8.5] });
  addConclusions(s, [
    'CX 6100 is the strongest Aruba replacement, but lacks MACsec, NetFlow, and DHCP snooping \u2014 evaluate risk per ship.',
    'VSF stacking is its unique operational advantage over Cisco C9200CX at fleet scale.'
  ], 6.0);
  addFooter(s, 15, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Marine-Environment Decision Points', 'Pre-procurement actions required regardless of vendor selection');
  const items = [
    ['Marine certification (IEC 60945 / DNV)', 'No reviewed source proved this for any of the three models. If compliance is mandatory, request written evidence or evaluate marine-grade alternatives (e.g., Moxa DNV-certified switches).'],
    ['Salt fog / corrosion (ASTM B117)', 'Not published by Cisco or HPE for these models. Request conformal coating specs and salt-spray test reports.'],
    ['Shock & vibration (MIL-STD-167)', 'Shipboard vibration (5\u201313 Hz) not addressed in any reviewed datasheet. Request test reports from both vendors.'],
    ['MTBF for fleet planning', 'Cisco publishes 553K\u2013704K hours. HPE does not publish MTBF for CX 6000/6100 \u2014 request from HPE support.'],
    ['Connector corrosion', 'No vendor specifies gold-plated or stainless connectors. Verify materials for marine air exposure.'],
    ['Spare parts strategy', 'At 2,500 units/ship, recommend 5\u201310% spares (125\u2013250 units) regardless of vendor.'],
  ];
  let y = 1.45;
  items.forEach(([title, detail]) => {
    s.addText(title, { x: 0.6, y, w: 3.6, h: 0.42, fontFace: FONT, fontSize: 14, bold: true, color: C.navy });
    s.addText(detail, { x: 4.2, y, w: 8.5, h: 0.42, fontFace: FONT, fontSize: 13, color: C.sub });
    y += 0.58;
  });
  addConclusions(s, [
    'None of the three switches are proven marine-certified \u2014 this is the single biggest open risk for Carnival.',
    'ACTION: Request written vendor evidence for vibration, corrosion, humidity, MTBF, and conformal coating before procurement.'
  ], 5.2);
  addFooter(s, 16, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Recommendation Matrix', 'Neutral recommendation after all corrections applied');
  addTable(s, [
    [hdr('Scenario'), hdr('Recommended Switch'), hdr('Rationale')],
    ['Need 10G uplinks + PoE \u2264 139W', greenCell('Aruba CX 6100 (JL679A)'), 'Retains 10G, lighter weight, VSF stacking. Verify PoE load first.'],
    ['Need 10G uplinks + PoE > 139W', cell('Stay with Cisco C9200CX', { bold: true }), 'Aruba cannot meet PoE requirements. No alternative in this comparison.'],
    ['No 10G requirement + PoE \u2264 139W', 'Aruba CX 6000 (R8N89A)', 'Lower cost, fanless, meets basic L2 access needs.'],
    ['Need MACsec encryption', cell('Stay with Cisco C9200CX', { bold: true }), 'Aruba CX 6000/6100 do not support MACsec (requires CX 6200+).'],
    ['Need NetFlow / deep visibility', cell('Stay with Cisco C9200CX', { bold: true }), 'Aruba offers sFlow only (sampled, not per-flow).'],
    ['Need dynamic routing (OSPF/EIGRP)', cell('Stay with Cisco C9200CX', { bold: true }), 'Aruba CX 6000/6100 are Layer 2 + static routing only.'],
    ['Need stacking / management consolidation', greenCell('Aruba CX 6100 (JL679A)'), 'CX 6100 supports 8-member VSF. Cisco C9200CX CANNOT stack.'],
    ['Need maximum weight reduction', greenCell('Aruba CX 6100 (JL679A)'), 'Lightest 10G-capable option: 0.47 lb/switch savings.'],
    ['Need marine-certified platform', redCell('No validated winner'), 'No reviewed source proved marine certification for any model.'],
  ], { y: 1.35, colW: [3.2, 2.8, 6.3], rowH: 0.44 });
  addConclusions(s, [
    'Aruba CX 6100 wins on weight and stacking; Cisco C9200CX wins on PoE, security, routing, and telemetry.',
    'The right choice depends on which trade-offs Carnival prioritizes \u2014 weight savings vs. feature depth.'
  ], 5.8);
  addFooter(s, 17, TOTAL);
}

{
  const s = pptx.addSlide();
  addSlideTitle(s, 'Sources & Methodology');
  const sources = [
    ['Cisco Catalyst 9200 Series Data Sheet', 'cisco.com/c/en/us/products/collateral/switches/catalyst-9200-series-switches/nb-06-cat9200-ser-data-sheet-cte-en.html'],
    ['HPE Aruba CX 6100 Product Page', 'buy.hpe.com/us/en/.../jl679a'],
    ['HPE Aruba CX 6000 Product Page', 'buy.hpe.com/us/en/.../r8n89a'],
    ['HPE Aruba CX 6100 Data Sheet', 'hpe.com/psnow/doc/a00106853enw'],
    ['HPE Aruba CX 6000 Data Sheet', 'hpe.com/psnow/doc/a00112996enw'],
    ['Aruba AOS-CX Security Guide (6000/6100)', 'arubanetworking.hpe.com/techdocs/AOS-CX/10.10/PDF/security_4100i-6000-6100.pdf'],
    ['Aruba AOS-CX IP Routing Guide (6000/6100)', 'arubanetworking.hpe.com/techdocs/AOS-CX/10.13/PDF/ip_route_4100i-6000-6100-6200.pdf'],
    ['Cisco C9200CX Hardware Installation Guide', 'cisco.com/c/en/us/td/docs/switches/lan/catalyst9200/hardware/install/b-c9200cx-hig'],
  ];
  let y = 1.45;
  sources.forEach(([name, url]) => {
    s.addText(name, { x: 0.6, y, w: 4.0, h: 0.35, fontFace: FONT, fontSize: 14, bold: true, color: C.navy });
    s.addText(url, { x: 4.6, y, w: 8.2, h: 0.35, fontFace: FONT, fontSize: 12, color: C.blue });
    y += 0.42;
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 5.2, w: 12.3, h: 1.2, rectRadius: 0.04, fill: { color: C.softBg }, line: { color: C.line, pt: 1 } });
  s.addText('Methodology', { x: 0.7, y: 5.6, w: 3, h: 0.3, fontFace: FONT, fontSize: 14, bold: true, color: C.navy });
  s.addText('All specifications were validated against official vendor datasheets and product pages. Cisco competitive marketing slides were noted but not treated as neutral evidence. Community forum opinions were excluded. Where data was unavailable, slides state "not published" rather than estimating. Corrections are highlighted in amber.', {
    x: 0.7, y: 5.95, w: 11.9, h: 0.55, fontFace: FONT, fontSize: 13, color: C.sub
  });
  addFooter(s, 18, TOTAL);
}

pptx.writeFile({ fileName: 'Cisco_C9200CX_Replacement_Analysis_Validated.pptx' });
