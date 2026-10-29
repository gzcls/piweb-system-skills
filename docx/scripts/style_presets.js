/** Reusable docx-js visual presets for generated documents. */
const STYLE_PRESETS = Object.freeze({
  contract: { bodyFont: 'Microsoft YaHei', headingFont: 'Microsoft YaHei', bodySize: 22, tableSize: 20, titleSize: 36, headingSize: 25, accent: '222222', tableHeaderFill: 'EEEEEE', tableBorder: 'AAAAAA', bodyAfter: 70, line: 280 },
  report: { bodyFont: 'Microsoft YaHei', headingFont: 'Microsoft YaHei', bodySize: 22, tableSize: 20, titleSize: 36, headingSize: 26, accent: '1F4E78', tableHeaderFill: 'DCE6F1', tableBorder: '9FBAD0', bodyAfter: 100, line: 300 },
  notice: { bodyFont: 'Microsoft YaHei', headingFont: 'Microsoft YaHei', bodySize: 22, tableSize: 20, titleSize: 34, headingSize: 25, accent: '333333', tableHeaderFill: 'F2F2F2', tableBorder: 'BFBFBF', bodyAfter: 60, line: 260 },
  branded: { bodyFont: 'Microsoft YaHei', headingFont: 'Microsoft YaHei', bodySize: 22, tableSize: 20, titleSize: 38, headingSize: 27, accent: '0B5CAD', tableHeaderFill: 'DCEBFA', tableBorder: '8CB3D9', bodyAfter: 90, line: 290 },
});
function createDocumentStyles(preset) {
  return {
    default: { document: { run: { font: preset.bodyFont, size: preset.bodySize }, paragraph: { spacing: { after: preset.bodyAfter, line: preset.line } } } },
    paragraphStyles: [{ id: 'Title', name: 'Title', run: { font: preset.headingFont, bold: true, color: preset.accent, size: preset.titleSize }, paragraph: { alignment: 'center', spacing: { after: preset.bodyAfter * 2 } } }, { id: 'Heading1', name: 'Heading 1', run: { font: preset.headingFont, bold: true, color: preset.accent, size: preset.headingSize }, paragraph: { spacing: { before: preset.bodyAfter * 2, after: preset.bodyAfter }, keepNext: true } }],
  };
}
function getStylePreset(name = 'contract', overrides = {}) {
  if (!STYLE_PRESETS[name]) throw new Error(`Unknown style preset: ${name}`);
  return { ...STYLE_PRESETS[name], ...overrides };
}
module.exports = { STYLE_PRESETS, getStylePreset, createDocumentStyles };
