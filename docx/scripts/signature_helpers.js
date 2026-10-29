/** Signature blocks that remain editable and do not depend on repeated underscores. */
const { BorderStyle, Paragraph, TableCell, TextRun, WidthType } = require('docx');
function signatureCell(party, width = 4653, options = {}) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: options.lineColor || 'AAAAAA' };
  return new TableCell({ width: { size: width, type: WidthType.DXA }, margins: { top: 180, bottom: 180, left: 120, right: 120 }, borders: { top: border, bottom: border, left: border, right: border }, children: [
    new Paragraph({ spacing: { after: 180 }, children: [new TextRun({ text: `${party}（盖章）`, bold: true })] }),
    new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: options.lineColor || 'AAAAAA' } }, spacing: { after: 180 }, children: [new TextRun({ text: '授权代表：' })] }),
    new Paragraph({ children: [new TextRun({ text: '签署日期：' })] }),
  ] });
}
module.exports = { signatureCell };
