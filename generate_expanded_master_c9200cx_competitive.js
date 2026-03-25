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
  if (title) slide.addText(title, { x: 0.72, y: 0.48, w: 11.4, h: 0.55, fontFace: FONT, fontSize: 27, color: C.blue });
  if (subtitle) slide.addText(subtitle, { x: 0.72, y: 1.0, w: 11.4, h: 0.38, fontFace: FONT, fontSize: 20, color: C.blue });
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
    fontSize: opts.fontSize ?? 10,
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
  slide.addText(text, { x: 0.8, y: 6.16, w: 11.7, h: 0.2, fontFace: FONT, fontSize: 11.8, color: C.white, bold: true, align: 'center' });
}

const WIDE = [1.55, 1.76, 1.76, 1.74, 1.76, 1.76, 1.77];

{
  const s = pptx.addSlide();
  s.background = { color: C.white };
  s.addText('Competitive Overview', { x: 0.72, y: 0.44, w: 6.2, h: 0.45, fontFace: FONT, fontSize: 16, color: C.dark });
  s.addText('C9200CX Competitive Comparison', { x: 0.72, y: 0.92, w: 8.8, h: 0.75, fontFace: FONT, fontSize: 28, color: C.blue });
  s.addText('Expanded master deck including Aruba, Arista, IE3500 Rugged, and IE3500H', { x: 0.72, y: 1.64, w: 9.6, h: 0.35, fontFace: FONT, fontSize: 16, color: C.text });
  s.addText('Prepared for: Carnival', { x: 0.72, y: 2.08, w: 4, h: 0.3, fontFace: FONT, fontSize: 15, color: C.text });
  s.addText('March 2026', { x: 0.72, y: 2.42, w: 2, h: 0.3, fontFace: FONT, fontSize: 15, color: C.text });
  s.addShape(pptx.ShapeType.line, { x: 0.72, y: 2.95, w: 11.8, h: 0, line: { color: C.line, pt: 1 } });
  bullets(s, [
    'Expanded from the validated Aruba, Arista, IE3500H, and official IE3500 Rugged fact set.',
    'Keeps C9200CX as the reference in all categories.',
    'Uses the same Cisco-style format as the prior master deck, but now includes IE3500 Rugged Series.'
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
    'IE3500 Rugged Series and IE3500H Heavy Duty split the industrial category: Rugged wins modular DIN-rail / RJ45 flexibility, while IE3500H wins sealed IP66/IP67 harsh-environment deployment.'
  ], 0.92, 1.6, 11.2, 17);
}

{
  const s = pptx.addSlide();
  hdr(s, 'Cisco Catalyst 9200CX Switches', 'Competitive Landscape', 3);
  addTable(s, [
    [h('Platform'), h('Compared Model / Series'), h('Role'), h('Primary Win')],
    [label('Cisco'), cell('C9200CX-12P-2X2G-E', C.gray), cell('Reference compact campus switch', C.gray), winner('All-around campus feature depth')],
    [label('Aruba'), cell('CX 6100 (JL679A)', C.white), cell('10G-capable Aruba compact alternative', C.white), good('Weight / VSF stacking')],
    [label('Aruba'), cell('CX 6000 (R8N89A)', C.gray), cell('Lower-end 1G-uplink option', C.gray), bad('Basic access only')],
    [label('Arista'), cell('CCS-710P-12', C.white), cell('Closest compact near-1:1 match', C.white), good('Closest clean hardware substitute')],
    [label('Cisco Industrial'), cell('IE3500 Rugged Series', C.gray), cell('DIN-rail modular industrial family', C.gray), good('RJ45 / modular / dual-DC industrial flexibility')],
    [label('Cisco Industrial'), cell('IE-3500H-12P2MU2XE', C.white), cell('Heavy-duty sealed industrial near-match', C.white), good('Rugged / IP67 / DC / M12')],
  ], { y: 1.65, colW: [1.55, 3.1, 4.05, 3.4], fontSize: 12.5, rowH: 0.54 });
}

{
  const s = pptx.addSlide();
  hdr(s, 'Executive Summary', 'Multi-vendor validated snapshot', 4);
  addTable(s, [
    [h('Metric'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('IE3500 Rugged'), h('Cisco IE3500H')],
    [label('Weight'), cell('6.6 lb', C.gray), winner('6.13 lb'), good('6.17 lb'), cell('not explicitly validated in final deck', C.gray), winner('3.90 lb representative SKU'), bad('11.40 lb')],
    [label('10G Uplinks'), winner('2x10G + 2x1G'), good('2x1/10G + 2x1G'), bad('No'), good('2x10G'), warn('1G or 10G SKU-dependent; not direct 12-port peer'), good('2x10G/1G')],
    [label('PoE Budget'), winner('240W'), bad('139W'), bad('139W'), warn('234W'), winner('240W base / up to 480W with expansion'), winner('240W')],
    [label('Stacking'), bad('No'), winner('VSF up to 8'), warn('Not confirmed'), bad('No evidence'), bad('No; modular expansion only'), bad('No')],
    [label('Environment'), bad('Campus / commercial'), bad('Campus / commercial'), bad('Campus / commercial'), bad('Campus / commercial'), good('Rugged industrial / IP30 / DIN rail'), winner('Rugged industrial / IP66-IP67')],
  ], { y: 1.65, colW: WIDE, fontSize: 9.3, rowH: 0.54 });
  footerBand(s, 'C9200CX wins campus breadth; Arista is the closest compact match; IE3500 Rugged and IE3500H split the industrial category');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Catalyst 9200CX Hardware Comparison', 'Core hardware and physical attributes', 5);
  addTable(s, [
    [h('Category'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('IE3500 Rugged'), h('Cisco IE3500H')],
    [label('CPU / Platform'), good('UADP 2.0 Mini / 4-core ARM'), cell('ARM Cortex-A9', C.gray), cell('ARM Cortex-A9', C.gray), good('Dual-core x86'), warn('Industrial modular platform; CPU not direct compare target'), warn('Industrial heavy-duty CPU not direct compare target')],
    [label('DRAM / Flash'), cell('4 GB / 8 GB', C.gray), good('4 GB / 16 GB'), good('4 GB / 16 GB corrected'), cell('4 GB / 8 GB', C.gray), winner('Series supports larger industrial memory / modular expansion context'), winner('8 GB / 5.1 GB user storage')],
    [label('Packet Buffer'), cell('6 MB', C.gray), winner('12.38 MB'), cell('not specified', C.gray), bad('2 MB'), cell('not explicitly normalized at series level', C.gray), cell('not officially published for exact SKU', C.gray)],
    [label('MAC Table'), winner('32K'), bad('8K'), bad('8K'), warn('Not clearly published'), cell('series-level values vary by SKU'), cell('24K', C.gray)],
    [label('Form Factor'), good('Compact fanless 1U'), good('Compact fanless 1U'), good('Compact fanless 1U'), good('Compact fanless 1U'), winner('DIN-rail modular industrial'), warn('Heavy-duty wall-mount / field form factor')],
    [label('Physical Fit'), good('9.6 in depth'), warn('~10 in class'), warn('~10 in class'), warn('9.8 in'), winner('Most compact industrial footprint'), bad('largest / heaviest rugged package')],
  ], { y: 1.62, colW: WIDE, fontSize: 8.8, rowH: 0.55 });
  footerBand(s, 'For compact campus hardware, Arista is the closest structural peer; Aruba leads packet buffer; IE3500 Rugged leads modular industrial compactness');
}

{
  const s = pptx.addSlide();
  hdr(s, 'PoE & Power Comparison', 'Critical for powered edge devices', 6);
  addTable(s, [
    [h('Category'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('IE3500 Rugged'), h('Cisco IE3500H')],
    [label('PoE Standard'), good('802.3at PoE+'), good('802.3at PoE+'), good('802.3at PoE+'), good('802.3at PoE+'), winner('Industrial PoE / 4PPoE capable with expansion options'), winner('PoE+ plus industrial 4PPoE-capable ports')],
    [label('PoE Budget'), winner('240W'), bad('139W'), bad('139W'), warn('234W'), winner('240W base / 360W or 480W with expansion'), winner('240W')],
    [label('Perpetual / Fast PoE'), winner('Yes'), bad('No'), bad('No'), warn('Persistent PoE on reboot'), cell('industrial power resilience emphasis', C.gray), cell('not the main differentiator', C.gray)],
    [label('Power Input'), cell('315W AC internal', C.gray), cell('Internal 165W', C.gray), cell('Internal', C.gray), warn('External AC PSU options'), winner('Dual DC 12–54V'), winner('12–54V DC industrial')],
    [label('Best for High-Power / OT endpoints'), good('Strong campus PoE headroom'), bad('Limited by 139W'), bad('Limited by 139W'), warn('Slightly lower total budget'), winner('Highest industrial PoE density / modular growth'), winner('Industrial high-power endpoint bias')],
  ], { y: 1.62, colW: WIDE, fontSize: 8.75, rowH: 0.56 });
  footerBand(s, 'Cisco wins campus PoE continuity; IE3500 Rugged wins modular industrial PoE density; IE3500H wins sealed DC industrial power architecture');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Security Features Comparison', '', 7);
  addTable(s, [
    [h('Feature'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('IE3500 Rugged'), h('Cisco IE3500H')],
    [label('MACsec'), winner('AES-256'), bad('No'), bad('No'), warn('Not fully substantiated'), winner('MACsec-256 at series level'), warn('MACsec-128 on -E')],
    [label('DHCP Snooping / DAI / IPSG'), winner('Yes / Yes / Yes'), bad('No / not confirmed / not confirmed'), bad('No / No / No'), warn('Not fully explicit in public source set'), good('Industrial security feature set documented'), warn('Industrial security focus, not direct campus parity')],
    [label('Trust / Secure Boot'), winner('Trust Anchor / Secure Boot'), cell('TPM integrated', C.gray), cell('TPM integrated', C.gray), warn('Security posture present, not equally explicit'), good('Cyber Vision + SEA with Network Advantage'), good('Secure boot / trust anchor class support')],
    [label('802.1X / TACACS / RADIUS'), good('Yes'), good('Yes'), good('Yes'), warn('Expected campus support'), good('Yes'), good('Yes')],
    [label('Best security category'), winner('Campus security depth'), bad('Limited'), bad('Basic'), warn('Good platform, less explicit public detail'), winner('Industrial security / OT visibility bundle'), good('Industrial security posture')],
  ], { y: 1.62, colW: WIDE, fontSize: 8.7, rowH: 0.58 });
  footerBand(s, 'Cisco is the campus-security winner; IE3500 Rugged adds the strongest explicit industrial security bundle; Aruba trails materially');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Software & Management Comparison', '', 8);
  addTable(s, [
    [h('Feature'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('IE3500 Rugged'), h('Cisco IE3500H')],
    [label('Operating System'), good('IOS XE'), good('AOS-CX'), good('AOS-CX'), good('EOS'), good('IOS XE industrial'), good('IOS XE industrial')],
    [label('Central Management'), good('Catalyst Center / Meraki'), good('Aruba Central'), good('Aruba Central'), good('CloudVision'), good('DNA Center support'), good('DNA Center support')],
    [label('Stacking / Consolidation'), bad('No'), winner('VSF up to 8'), warn('Not confirmed'), bad('No evidence'), bad('No; modular expansion only'), bad('No')],
    [label('Programmability'), winner('RESTCONF / NETCONF / YANG / gNMI'), bad('REST API'), bad('REST API'), good('EOS APIs / programmability'), good('Cisco IOS XE automation family'), good('Cisco IOS XE automation family')],
    [label('Telemetry'), winner('Model-driven streaming telemetry'), bad('SNMP + sFlow'), bad('SNMP + sFlow'), good('CloudVision telemetry'), good('Industrial Cisco telemetry / OT visibility'), good('Industrial IOS XE / DNA support')],
    [label('App Visibility'), winner('NBAR2 / AVC'), bad('No'), bad('No'), bad('No NBAR equivalent'), warn('OT-focused visibility, different model'), bad('OT focus, not campus app vis')],
  ], { y: 1.62, colW: WIDE, fontSize: 8.55, rowH: 0.56 });
  footerBand(s, 'Aruba wins stacking; Arista wins EOS/CloudVision operations; Cisco wins campus programmability and app visibility; Rugged adds OT-oriented management value');
}

{
  const s = pptx.addSlide();
  hdr(s, 'L2 / L3 Protocol Support', '', 9);
  addTable(s, [
    [h('Protocol / Role'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('IE3500 Rugged'), h('Cisco IE3500H')],
    [label('Static Routing'), good('Yes'), good('Yes'), good('Yes'), warn('Expected but not fully enumerated'), good('Yes'), good('Yes')],
    [label('OSPF / EIGRP / IS-IS / RIP'), winner('Explicitly documented'), bad('No / L2 only bias'), bad('No / L2 only bias'), warn('Not fully enumerated publicly'), good('Industrial IOS XE family support'), good('Industrial IOS XE family support')],
    [label('BGP / EVPN'), good('Basic BGP on C9200CX'), bad('No'), bad('No'), good('EVPN segmentation positioning'), warn('Series-level industrial routing / not positioned as EVPN compact peer'), winner('BGP supported in reviewed IE set')],
    [label('PTP / Timing'), winner('IEEE 1588v2'), bad('Not a key differentiator'), bad('Not a key differentiator'), warn('Not directly substantiated'), winner('TSN / frame-preemption / industrial timing posture'), winner('IEEE 1588v2')],
    [label('Industrial Protocol Relevance'), bad('Campus platform'), bad('Campus platform'), bad('Campus platform'), bad('Campus platform'), winner('Industrial redundancy / OT protocols / TSN posture'), winner('OT / industrial protocol alignment')],
  ], { y: 1.62, colW: WIDE, fontSize: 8.55, rowH: 0.58 });
  footerBand(s, 'Cisco wins documented campus routing depth; IE3500 Rugged and IE3500H dominate OT protocol and timing relevance');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Reliability & Environmental Specifications', '', 10);
  addTable(s, [
    [h('Criterion'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('IE3500 Rugged'), h('Cisco IE3500H')],
    [label('MTBF'), winner('Published 553K–704K hrs range'), warn('Not published'), warn('Not published'), warn('Not validated in deck'), warn('Not primary published takeaway in source set'), warn('Not the primary published takeaway in deck')],
    [label('Operating Temp'), cell('-5°C to 45°C', C.gray), cell('0°C to 45°C', C.gray), cell('0°C to 45°C', C.gray), cell('campus compact class', C.gray), winner('-40°C to 75°C'), winner('-40°C to 75°C')],
    [label('Humidity / Marine Margin'), warn('5–90% non-condensing'), warn('5–90%'), warn('5–90%'), warn('Not positioned for marine extremes'), good('Best protected industrial cabinet posture'), winner('Best sealed harsh-environment posture')],
    [label('Ingress / Shock / Vibration'), bad('Not rugged rated'), bad('Not rugged rated'), bad('Not rugged rated'), bad('Not rugged rated'), good('Shock / vibration / surge hardened, IP30'), winner('Industrial hardening / IP66-IP67')],
    [label('Marine Readiness'), bad('Not proven marine-certified'), bad('Not proven marine-certified'), bad('Not proven marine-certified'), bad('Not proven marine-certified'), good('Good for protected industrial / cabinet use'), winner('Closest practical harsh-environment fit')],
  ], { y: 1.62, colW: WIDE, fontSize: 8.55, rowH: 0.58 });
  footerBand(s, 'C9200CX is the only platform with clearly cited MTBF here; IE3500 Rugged and IE3500H are the real environmental contenders');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Telemetry & Monitoring Capabilities', '', 11);
  addTable(s, [
    [h('Feature'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('IE3500 Rugged'), h('Cisco IE3500H')],
    [label('SNMP / Syslog / LLDP'), good('Yes'), good('Yes'), good('Yes'), good('Yes'), good('Yes'), good('Yes')],
    [label('Flow Monitoring'), winner('Flexible NetFlow'), bad('sFlow only'), bad('sFlow only'), warn('sFlow'), warn('Cisco OT visibility / monitoring stack'), warn('DNA support but not NetFlow-centric in deck')],
    [label('Streaming Telemetry'), winner('Model-driven gRPC push'), bad('SNMP pull bias'), bad('SNMP pull bias'), good('CloudVision telemetry'), good('Industrial Cisco telemetry / OT visibility'), good('Industrial Cisco automation family')],
    [label('YANG / Open APIs'), winner('OpenConfig + native YANG'), bad('No'), bad('No'), good('EOS APIs'), good('IOS XE automation'), good('IOS XE automation')],
    [label('Best observability category'), winner('Cisco campus'), bad('Basic'), bad('Basic'), good('Strong Arista telemetry ops'), good('Strong OT visibility posture'), warn('Operational OT visibility, different goal')],
  ], { y: 1.62, colW: WIDE, fontSize: 8.55, rowH: 0.58 });
  footerBand(s, 'Cisco wins campus observability depth; Arista is the strongest compact alternative; IE3500 Rugged adds OT-oriented monitoring value');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Stacking, Scale, and Weight Impact', '', 12);
  addTable(s, [
    [h('Aspect'), h('Cisco C9200CX'), h('Aruba CX 6100'), h('Aruba CX 6000'), h('Arista 710P-12'), h('IE3500 Rugged'), h('Cisco IE3500H')],
    [label('Stacking Support'), bad('No'), winner('VSF up to 8'), warn('Not confirmed'), bad('No evidence in deck'), bad('No; modular expansion only'), bad('No')],
    [label('Mgmt entities at 2,500'), bad('2,500 units'), winner('~313 VSF stacks'), bad('2,500 assumed'), bad('2,500 units'), warn('Modular growth reduces node count per cabinet, not stacking'), bad('2,500 rugged nodes')],
    [label('Weight Impact'), cell('Baseline 6.6 lb', C.gray), winner('Best validated reduction'), good('Moderate reduction'), warn('Not explicit in final deck'), winner('Representative rugged SKU much lighter than IE3500H'), bad('Much heavier')],
    [label('Fleet-scale Ops Winner'), bad('No stacking'), winner('Best management consolidation'), bad('Too basic'), warn('Closest compact alt, no stacking edge'), good('Best modular industrial cabinet scaling'), bad('Environment-focused, not fleet simplification')],
  ], { y: 1.75, colW: WIDE, fontSize: 8.7, rowH: 0.62 });
  footerBand(s, 'Aruba CX 6100 is the fleet-scale weight/stacking winner; IE3500 Rugged is the better modular industrial scaling story; Cisco wins feature depth');
}

{
  const s = pptx.addSlide();
  hdr(s, 'Recommendation Matrix', '', 13);
  addTable(s, [
    [h('Priority'), h('Winner'), h('Reason')],
    [label('Best all-around campus access'), winner('Cisco C9200CX'), cell('PoE, security, routing, telemetry, app visibility, and Cisco campus integration.', C.gray)],
    [label('Closest clean compact substitute'), winner('Arista CCS-710P-12'), cell('Best 1:1 compact role alignment with fanless 12-port PoE and 2x10G uplinks.', C.white)],
    [label('Best protected industrial / cabinet deployment fit'), winner('Cisco IE3500 Rugged Series'), cell('DIN-rail modular family, dual DC, RJ45, industrial hardening, TSN / OT posture.', C.gray)],
    [label('Best sealed harsh-environment / field deployment fit'), winner('Cisco IE-3500H-12P2MU2XE'), cell('IP66/IP67, DC, M12, industrial hardening, harsh-environment suitability.', C.white)],
    [label('Best weight / stacking compromise'), winner('Aruba CX 6100'), cell('Lowest validated weight and VSF stacking, but weaker security and telemetry.', C.gray)],
    [label('Avoid when 10G / security / visibility matter'), bad('Aruba CX 6000'), cell('Basic L2 access only; most constrained option in this study.', C.white)],
  ], { y: 1.7, colW: [3.15, 3.2, 5.75], fontSize: 11.8, rowH: 0.6 });
}

{
  const s = pptx.addSlide();
  hdr(s, 'Sources & Methodology', '', 14);
  bullets(s, [
    'Source decks merged: Cisco_C9200CX_Replacement_Analysis_Validated.pptx, C9200CX_vs_Arista_710P_Validated.pptx, C9200CX_vs_IE3500H_Validated.pptx.',
    'IE3500 Rugged Series was added from an official Cisco series-level fact set and is shown conservatively as a series column, not a forced 12-port 1:1 SKU match.',
    'Where one vendor had less explicit public detail than Cisco, the merged deck keeps the claim conservative instead of guessing.',
    'Winners are category-specific, not absolute; different platforms win different deployment conditions.'
  ], 0.92, 1.5, 11.2, 16);
  addTable(s, [
    [h('Source Set'), h('Use In This Deck')],
    [label('Aruba validated deck'), cell('Used for weight, stacking, power, security, management, telemetry, recommendation logic', C.gray)],
    [label('Arista validated deck'), cell('Used for 1:1 compact equivalence, efficiency, EOS / CloudVision, caveats', C.white)],
    [label('IE3500H validated deck'), cell('Used for ruggedness, industrial certifications, environmental fit, OT positioning', C.gray)],
    [label('IE3500 Rugged official fact set'), cell('Used for DIN-rail modular industrial positioning, dual-DC power, RJ45, TSN, and industrial comparison points', C.white)],
  ], { y: 4.45, colW: [2.95, 8.95], fontSize: 11.4, rowH: 0.54 });
}

pptx.writeFile({ fileName: 'expanded_master_c9200cx_competitive_comparison.pptx' });
