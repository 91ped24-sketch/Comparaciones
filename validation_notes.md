# Validated switch replacement analysis

This folder includes a corrected PPTX built from reviewed official Cisco and HPE Aruba sources.

## Main corrections versus the parsed deck

1. **Cisco C9200CX does not support StackWise stacking**.
   - Cisco 9200 datasheet explicitly says stacking is not available on C9200CX switches.
2. **Aruba CX 6100 does support VSF stacking**.
   - HPE store page for the CX 6100 series states 8-member VSF support.
3. **Competitive marketing slides are not neutral proof**.
   - They can inform positioning, but they should not be treated as validated specifications.
4. **Marine suitability is not proven by the reviewed product pages**.
   - No reviewed source proved IEC 60945 / DNV certification for the exact models.

## Deliverable produced

- `validated_switch_replacement_analysis.pptx`

## Evidence set used

- Cisco Catalyst 9200 Series Data Sheet
- HPE Aruba Networking CX 6100 JL679A product page
- HPE Aruba Networking CX 6000 R8N89A product page
- HPE Aruba Networking CX 6100 Switch Series data sheet
- HPE Aruba Networking CX 6000 Switch Series data sheet
