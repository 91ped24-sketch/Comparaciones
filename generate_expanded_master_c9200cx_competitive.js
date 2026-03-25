const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Carnival Network Engineering';
pptx.company = 'Carnival Corporation';
pptx.subject = 'Expanded merged C9200CX competitive comparison';
pptx.title = 'Expanded C9200CX Competitive Comparison';
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
const TOTAL = 14;

function hdr(slide, title, subtitle, page) {
  slide.background = { color: C.white };
  slide.addText(title, { x: 0.72, y: 0.48, w: 11.3, h: 0.55, fontFace: FONT, fontSize: 27, color: C.blue });
  if (subtitle) slide.addText(subtitle, { x: 0.72, y: 1.0, w: 11.3, h: 0.38, fontFace: FONT, fontSize: 20, color: C.blue });
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
    fontSize: opts.fontSize ?? 11,
    color: C.text,
    border: { pt: 0.5, color: 'D9D9DD' },
    margin: 0.05,
    valign: 'mid',
    autoPage: false,
  });
}

function h(text) { return { text, options: { bold: true, color: C.dark, fill: { color: C.white }, align: 'center' } }; }
function label(text) { return { text, options: { bold: true, fill: { color: C.white }, color: C.dark } }; }
function cell(text, fill = C.gray2, extra = {}) { return { text, options: { fill: { color: fill }, color: C.text, align: 'center', ...extra } }; }
function winner(text) { return { text: `🏆 ${text}`, options: { fill: { color: C.gray2 }, color: C.green, bold: true, align: 'center' } }; }
function good(text) { return { text, options: { fill: { color: C.gray2 }, color: C.green, bold: true, align: 'center' } }; }
function warn(text) { return { text, options: { fill: { color: C.gray2 }, color: C.amber, bold: true, align: 'center' } }; }
function bad(text) { return { text, options: { fill: { color: C.gray2 }, color: C.red, bold: true, align: 'center' } }; }

function bullets(slide, items, x, y, w, fs = 15) {
  let yy = y;
  for (const item of items) {
    slide.addText([{ text: '• ', options: { color: C.text } }, { text: item, options: { color: C.text } }], {
      x, y: yy, w, h: 0.5, fontFace: FONT, fontSize: fs, breakLine: false, margin: 0
    });
    yy += 0.66;
  }
}

function footerBand(slide, text) {
  slide.addShape(pptx.ShapeType.rect, { x: 0.6, y: 6.08, w: 12.1, h: 0.38, fill: { color: C.blue }, line: { color: C.blue, pt: 0 } });
  slide.addText(text, { x: 0.8, y: 6.16, w: 11.7, h: 0.2, fontFace: FONT, fontSize: 12.5, color: C.white, bold: true, align: 'center' });
}

{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  s.addText('Competitive Overview', { x: 0.72, y: 0.44, w: 6.2, h: 0.45, fontFace: FONT, fontSize: 16, color: C.dark });
  s.addText('C9200CX Competitive Comparison', { x: 0.72, y: 0.92, w: 8.8, h: 0.75, fontFace: FONT, fontSize: 28, color: C.blue });
  s.addText('Expanded master deck based on the complete Aruba replacement analysis structure', { x: 0.72, y: 1.64, w: 8.9, h: 0.35, fontFace: FONT, fontSize: 16, color: C.text });
  s.addText('Prepared for: Carnival', { x: 0.72, y: 2.08, w: 4, h: 0.3, fontFace: FONT, fontSize: 15, color: C.text });
  s.addText('March 2026', { x: 0.72, y: 2.42, w: 2, h: 0.3, fontFace: FONT, fontSize: 15, color: C.text });
  s.addShape(pptx.ShapeType.line, { x: 0.72, y: 2.95, w: 11.8, h: 0, line: { color: C.line, pt: 1 } });
  bullets(s, [
    'Expanded from the validated Aruba, Arista, and IE3500H decks.',
    'Keeps C9200CX as the reference in all categories.',
    'Uses the Cisco-style format from the prior master deck, but with fuller topic coverage.'
  ], 0.92, 3.45, 11.0, 18);
  hdr(s, '', '', 1);
}

{
  const s = pptx.addSlide();
  hdr(s, 'Key Technical Findings:', '', 2);
  bullets(s, [
    'C9200CX remains the best all-around campus access platform when PoE headroom, MACsec-256, telemetry, visibility, and Cisco campus integration matter most.',
    'Arista CCS-710P-12 is the closest true compact near-match, with very similar physical role and lower switching power draw.',
    'Aruba CX 6100 wins the weight and stacking category, but it trades away major security, telemetry, and PoE headroom.',
    'IE-3500H-12P2MU2XE is the environment winner for rugged / marine-adjacent / OT deployment, not a direct campus peer.'
  ], 0.92, 1.6, 11.2, 17);
}

{
  const s = pptx.addSlide();
  hdr(s, 'Cisco Catalyst 9200CX Switches', 'Competitive Landscape', 3);
  addTable(s, [
    [h('Platform'), h('Compared Model'), h('Role'), h('Primary Win')],
    [label('Cisco'), cell('C9200CX-12P-2X2G-E', C.gray), cell('Reference compact campus switch', C.gray), winner('All-around campus feature depth')],
    [label('Aruba'), cell('CX 6100 (JL679A)', C.white), cell('10G-capable Aruba compact alternative', C.white), good('Weight / VSF stacking')],
    [label('Aruba'), cell('CX 6000 (R8N89A)', C.gray), cell('Lower-end 1G-uplink option', C.gray), bad('Basic access only')],
    [label('Arista'), cell('CCS-710P-12', C.white), cell('Closest compact near-1:1 match', C.white), good('Closest clean hardware substitute')],
    [label('Cisco Industrial'), cell('IE-3500H-12P2MU2XE', C.gray), cell('Industrial / harsh-environment near-match', C.gray), good('Rugged / IP67 / DC / M12')],
  ], { y: 1.7, colW: [1.6, 3.4, 3.7, 3.4], fontSize: 13, rowH: 0.58 });
}

{
  const s = pptx.addSlide();
  hdr(s, 'Executive Summary', 'Multi-vendor validated snapshot', 4);
  addTable(s, [
    [h('Metric'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('Weight'), cell('6.6 lb', C.gray), winner('6.13 lb'), good('6.17 lb'), cell('not explicitly validated in final deck', C.gray), bad('11.40 lb')],
    [label('10G Uplinks'), winner('2x10G + 2x1G'), good('2x1/10G + 2x1G'), bad('No'), good('2x10G'), good('2x10G/1G')],
    [label('PoE Budget'), winner('240W'), bad('139W'), bad('139W'), warn('234W'), winner('240W')],
    [label('Stacking'), bad('No'), winner('VSF up to 8'), warn('Not confirmed'), bad('No evidence'), bad('No')],
    [label('Environment'), bad('Campus / commercial'), bad('Campus / commercial'), bad('Campus / commercial'), bad('Campus / commercial'), winner('Rugged industrial / IP66/IP67')],
  ], { y: 1.7, colW: [1.7, 2.05, 2.05, 2.0, 2.0, 2.15], fontSize: 11.4, rowH: 0.56 });
  footerBand(s, 'C9200CX wins the broad campus category; Arista is the closest compact match; IE3500H wins the environment category');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Catalyst 9200CX Hardware Comparison', 'Core hardware and physical attributes', 5);
  addTable(s, [
    [h('Category'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('CPU / Platform'), good('UADP 2.0 Mini / 4-core ARM'), cell('ARM Cortex-A9', C.gray), cell('ARM Cortex-A9', C.gray), good('Dual-core x86'), warn('Industrial platform CPU not core comparison target')],
    [label('DRAM / Flash'), cell('4 GB / 8 GB', C.gray), good('4 GB / 16 GB'), good('4 GB / 16 GB corrected'), cell('4 GB / 8 GB', C.gray), winner('8 GB / 5.1 GB user storage')],
    [label('Packet Buffer'), cell('6 MB', C.gray), winner('12.38 MB'), cell('not specified', C.gray), bad('2 MB'), cell('not officially published', C.gray)],
    [label('MAC Table'), winner('32K'), bad('8K'), bad('8K'), warn('Not clearly published'), cell('24K', C.gray)],
    [label('Form Factor'), good('Compact fanless 1U'), good('Compact fanless 1U'), good('Compact fanless 1U'), good('Compact fanless 1U'), warn('Heavy-duty wall-mount / field form factor')],
    [label('Depth / Physical Fit'), good('9.6 in'), warn('~10 in class'), warn('~10 in class'), warn('9.8 in'), bad('larger / heavier rugged package')],
  ], { y: 1.7, colW: [1.75, 2.08, 2.03, 2.0, 2.0, 2.14], fontSize: 10.8, rowH: 0.55 });
  footerBand(s, 'For compact campus hardware, Arista is the closest structural peer; Aruba leads packet buffer; Cisco leads documented L2 scale');
}

{
  const s = pptx.addSlide();
  hdr(s, 'PoE & Power Comparison', 'Critical for powered edge devices', 6);
  addTable(s, [
    [h('Category'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('PoE Standard'), good('802.3at PoE+'), good('802.3at PoE+'), good('802.3at PoE+'), good('802.3at PoE+'), winner('PoE+ plus industrial 4PPoE-capable ports')],
    [label('PoE Budget'), winner('240W'), bad('139W'), bad('139W'), warn('234W'), winner('240W')],
    [label('Perpetual / Fast PoE'), winner('Yes'), bad('No'), bad('No'), warn('Persistent PoE on reboot'), cell('not the main differentiator', C.gray)],
    [label('Power Input'), cell('315W AC internal', C.gray), cell('Internal 165W', C.gray), cell('Internal', C.gray), warn('External AC PSU options'), winner('12-54V DC industrial')],
    [label('Best for High-Power / OT endpoints'), good('Strong campus PoE headroom'), bad('Limited by 139W'), bad('Limited by 139W'), warn('Slightly lower total budget'), winner('Industrial high-power endpoint bias')],
  ], { y: 1.7, colW: [1.8, 2.04, 2.02, 1.98, 2.02, 2.24], fontSize: 10.8, rowH: 0.56 });
  footerBand(s, 'Cisco wins campus PoE continuity; IE3500H wins DC / industrial power architecture; Aruba is gated by 139W');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Security Features Comparison', '', 7);
  addTable(s, [
    [h('Feature'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('MACsec'), winner('AES-256'), bad('No'), bad('No'), warn('Not fully substantiated'), warn('MACsec-128 on -E')],
    [label('DHCP Snooping / DAI / IPSG'), winner('Yes / Yes / Yes'), bad('No / not confirmed / not confirmed'), bad('No / No / No'), warn('Not fully explicit in public source set'), warn('Industrial security focus, not direct campus parity')],
    [label('Trust / Secure Boot'), winner('Trust Anchor / Secure Boot'), cell('TPM integrated', C.gray), cell('TPM integrated', C.gray), warn('Security posture present, not equally explicit'), good('Secure boot / trust anchor class support')],
    [label('802.1X / TACACS / RADIUS'), good('Yes'), good('Yes'), good('Yes'), warn('Expected campus support'), good('Yes')],
    [label('Best security category'), winner('Campus security depth'), bad('Limited'), bad('Basic'), warn('Good platform, less explicit public detail'), good('Industrial security posture')],
  ], { y: 1.7, colW: [1.95, 2.0, 2.0, 1.95, 2.0, 2.2], fontSize: 10.4, rowH: 0.58 });
  footerBand(s, 'Cisco is the clear campus-security winner; IE3500H is the strongest rugged-security option; Aruba trails materially');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Software & Management Comparison', '', 8);
  addTable(s, [
    [h('Feature'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('Operating System'), good('IOS XE'), good('AOS-CX'), good('AOS-CX'), good('EOS'), good('IOS XE industrial')],
    [label('Central Management'), good('Catalyst Center / Meraki'), good('Aruba Central'), good('Aruba Central'), good('CloudVision'), good('DNA Center support')],
    [label('Stacking / Consolidation'), bad('No'), winner('VSF up to 8'), warn('Not confirmed'), bad('No evidence'), bad('No')],
    [label('Programmability'), winner('RESTCONF / NETCONF / YANG / gNMI'), bad('REST API'), bad('REST API'), good('EOS APIs / programmability'), good('Cisco IOS XE automation family')],
    [label('Telemetry'), winner('Model-driven streaming telemetry'), bad('SNMP + sFlow'), bad('SNMP + sFlow'), good('CloudVision telemetry'), good('Industrial IOS XE / DNA support')],
    [label('App Visibility'), winner('NBAR2 / AVC'), bad('No'), bad('No'), bad('No NBAR equivalent'), bad('OT focus, not campus app vis')],
  ], { y: 1.7, colW: [1.9, 2.03, 1.98, 1.94, 2.05, 2.2], fontSize: 10.2, rowH: 0.56 });
  footerBand(s, 'Aruba wins stacking; Arista wins EOS/CloudVision operations; Cisco wins programmability, telemetry, and app visibility');
}

{
  const s = pptx.addSlide();
  hdr(s, 'L2 / L3 Protocol Support', '', 9);
  addTable(s, [
    [h('Protocol / Role'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('Static Routing'), good('Yes'), good('Yes'), good('Yes'), warn('Expected but not fully enumerated'), good('Yes')],
    [label('OSPF / EIGRP / IS-IS / RIP'), winner('Explicitly documented'), bad('No / L2 only bias'), bad('No / L2 only bias'), warn('Not fully enumerated publicly'), good('Yes industrial IOS XE set')],
    [label('BGP / EVPN'), good('Basic BGP on C9200CX'), bad('No'), bad('No'), good('EVPN segmentation positioning'), winner('BGP supported in reviewed IE set')],
    [label('PTP / Timing'), winner('IEEE 1588v2'), bad('Not a key differentiator'), bad('Not a key differentiator'), warn('Not directly substantiated'), winner('IEEE 1588v2')],
    [label('Industrial Protocol Relevance'), bad('Campus platform'), bad('Campus platform'), bad('Campus platform'), bad('Campus platform'), winner('OT / industrial protocol alignment')],
  ], { y: 1.72, colW: [2.1, 1.95, 1.95, 1.9, 2.0, 2.2], fontSize: 10.2, rowH: 0.58 });
  footerBand(s, 'Cisco wins documented campus routing depth; IE3500H wins industrial protocol alignment; Aruba is the most limited L3 path');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Reliability & Environmental Specifications', '', 10);
  addTable(s, [
    [h('Criterion'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('MTBF'), winner('Published 553K–704K hrs range'), warn('Not published'), warn('Not published'), warn('Not validated in deck'), warn('Not the primary published takeaway in deck')],
    [label('Operating Temp'), cell('-5°C to 45°C', C.gray), cell('0°C to 45°C', C.gray), cell('0°C to 45°C', C.gray), cell('campus compact class', C.gray), winner('-40°C to 75°C')],
    [label('Humidity / Marine Margin'), warn('5–90% non-condensing'), warn('5–90%'), warn('5–90%'), warn('Not positioned for marine extremes'), winner('Best harsh-environment posture')],
    [label('Ingress / Shock / Vibration'), bad('Not rugged rated'), bad('Not rugged rated'), bad('Not rugged rated'), bad('Not rugged rated'), winner('Industrial hardening / IP66-IP67')],
    [label('Marine Readiness'), bad('Not proven marine-certified'), bad('Not proven marine-certified'), bad('Not proven marine-certified'), bad('Not proven marine-certified'), good('Closest practical harsh-environment fit')],
  ], { y: 1.7, colW: [1.95, 2.0, 1.95, 1.9, 2.0, 2.3], fontSize: 10.2, rowH: 0.58 });
  footerBand(s, 'C9200CX is the only platform with clearly cited MTBF in this source set; IE3500H is the only true environmental winner');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Telemetry & Monitoring Capabilities', '', 11);
  addTable(s, [
    [h('Feature'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('SNMP / Syslog / LLDP'), good('Yes'), good('Yes'), good('Yes'), good('Yes'), good('Yes')],
    [label('Flow Monitoring'), winner('Flexible NetFlow'), bad('sFlow only'), bad('sFlow only'), warn('sFlow'), warn('DNA support but not NetFlow-centric in deck')],
    [label('Streaming Telemetry'), winner('Model-driven gRPC push'), bad('SNMP pull bias'), bad('SNMP pull bias'), good('CloudVision telemetry'), good('Industrial Cisco automation family')],
    [label('YANG / Open APIs'), winner('OpenConfig + native YANG'), bad('No'), bad('No'), good('EOS APIs'), good('IOS XE automation')],
    [label('Best observability category'), winner('Cisco'), bad('Basic'), bad('Basic'), good('Strong Arista telemetry ops'), warn('Operational OT visibility, different goal')],
  ], { y: 1.7, colW: [1.95, 2.0, 1.95, 1.9, 2.0, 2.3], fontSize: 10.2, rowH: 0.58 });
  footerBand(s, 'Cisco wins observability depth; Arista is the strongest alternative for telemetry-centric operations');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Stacking, Scale, and Weight Impact', '', 12);
  addTable(s, [
    [h('Aspect'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('Cisco IE3500H')],
    [label('Stacking Support'), bad('No'), winner('VSF up to 8'), warn('Not confirmed'), bad('No evidence in deck'), bad('No')],
    [label('Mgmt entities at 2,500'), bad('2,500 units'), winner('~313 VSF stacks'), bad('2,500 assumed'), bad('2,500 units'), bad('2,500 rugged nodes')],
    [label('Weight Impact'), cell('Baseline 6.6 lb', C.gray), winner('Best validated reduction'), good('Moderate reduction'), warn('Not explicit in final deck'), bad('Much heavier')],
    [label('Fleet-scale Ops Winner'), bad('No stacking'), winner('Best management consolidation'), bad('Too basic'), warn('Closest compact alt, no stacking edge'), bad('Environment-focused, not fleet simplification')],
  ], { y: 1.8, colW: [2.15, 1.95, 1.95, 1.9, 2.0, 2.15], fontSize: 10.3, rowH: 0.62 });
  footerBand(s, 'Aruba CX 6100 is the clear fleet-scale weight/stacking winner; Cisco wins where feature depth matters more than management consolidation');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Recommendation Matrix', '', 13);
  addTable(s, [
    [h('Priority'), h('Winner'), h('Reason')],
    [label('Best all-around campus access'), winner('Cisco C9200CX'), cell('PoE, security, routing, telemetry, app visibility, and Cisco campus integration.', C.gray)],
    [label('Closest clean compact substitute'), winner('Arista CCS-710P-12'), cell('Best 1:1 compact role alignment with fanless 12-port PoE and 2x10G uplinks.', C.white)],
    [label('Best rugged / marine-adjacent deployment fit'), winner('Cisco IE-3500H-12P2MU2XE'), cell('IP66/IP67, DC, M12, industrial hardening, harsh-environment suitability.', C.gray)],
    [label('Best weight / stacking compromise'), winner('Aruba CX 6100'), cell('Lowest validated weight and VSF stacking, but weaker security and telemetry.', C.white)],
    [label('Avoid when 10G / security / visibility matter'), bad('Aruba CX 6000'), cell('Basic L2 access only; most constrained option in this study.', C.gray)],
  ], { y: 1.78, colW: [3.15, 3.25, 5.7], fontSize: 12.4, rowH: 0.64 });
}

{
  const s = pptx.addSlide();
  hdr(s, 'Sources & Methodology', '', 14);
  bullets(s, [
    'Source decks merged: Cisco_C9200CX_Replacement_Analysis_Validated.pptx, C9200CX_vs_Arista_710P_Validated.pptx, C9200CX_vs_IE3500H_Validated.pptx.',
    'This expanded master deck inherits the validated findings from those generated decks and restructures them into a Cisco-style competitive format.',
    'Where one vendor had less explicit public detail than Cisco, the merged deck keeps the claim conservative instead of guessing.',
    'Winners are category-specific, not absolute; different platforms win different deployment conditions.'
  ], 0.92, 1.6, 11.2, 17);
  addTable(s, [
    [h('Source Set'), h('Use In This Deck')],
    [label('Aruba validated deck'), cell('Used for weight, stacking, power, security, management, telemetry, recommendation logic', C.gray)],
    [label('Arista validated deck'), cell('Used for 1:1 compact equivalence, efficiency, EOS / CloudVision, caveats', C.white)],
    [label('IE3500H validated deck'), cell('Used for ruggedness, industrial certifications, environmental fit, OT positioning', C.gray)],
  ], { y: 4.6, colW: [2.8, 9.1], fontSize: 12, rowH: 0.58 });
}

pptx.writeFile({ fileName: 'expanded_master_c9200cx_competitive_comparison.pptx' });
