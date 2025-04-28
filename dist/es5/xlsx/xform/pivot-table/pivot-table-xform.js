"use strict";

const XmlStream = require('../../../utils/xml-stream');
const BaseXform = require('../base-xform');
class PivotTableXform extends BaseXform {
  constructor() {
    super();
    this.map = {};
  }
  prepare(model) {
    // TK
  }
  get tag() {
    // http://www.datypic.com/sc/ooxml/e-ssml_pivotTableDefinition.html
    return 'pivotTableDefinition';
  }
  render(xmlStream, model) {
    // eslint-disable-next-line no-unused-vars
    const {
      rows,
      columns,
      values,
      metric,
      cacheFields,
      cacheId
    } = model;

    // Examples
    // --------
    // rows: [0, 1], // only 2 items possible for now
    // columns: [2], // only 1 item possible for now
    // values: [4], // only 1 item possible for now
    // metric: 'sum', // only 'sum' possible for now
    //
    // the numbers are indices into `cacheFields`.

    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...PivotTableXform.PIVOT_TABLE_ATTRIBUTES,
      'xr:uid': '{267EE50F-B116-784D-8DC2-BA77DE3F4F4A}',
      name: 'PivotTable2',
      cacheId,
      applyNumberFormats: '0',
      applyBorderFormats: '0',
      applyFontFormats: '0',
      applyPatternFormats: '0',
      applyAlignmentFormats: '0',
      applyWidthHeightFormats: '1',
      dataCaption: 'Values',
      updatedVersion: '8',
      minRefreshableVersion: '3',
      useAutoFormatting: '1',
      itemPrintTitles: '1',
      createdVersion: '8',
      indent: '0',
      compact: '0',
      compactData: '0',
      multipleFieldFilters: '0'
    });

    // Note: keeping this pretty-printed and verbose for now to ease debugging.
    //
    // location: ref="A3:E15"
    // pivotFields
    // rowFields and rowItems
    // colFields and colItems
    // dataFields
    // pivotTableStyleInfo
    xmlStream.writeXml("\n      <location ref=\"A3:E15\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\" />\n      <pivotFields count=\"".concat(cacheFields.length, "\">\n        ").concat(renderPivotFields(model), "\n      </pivotFields>\n      <rowFields count=\"").concat(rows.length, "\">\n        ").concat(rows.map(rowIndex => "<field x=\"".concat(rowIndex, "\" />")).join('\n    '), "\n      </rowFields>\n      <rowItems count=\"1\">\n        <i t=\"grand\"><x /></i>\n      </rowItems>\n      <colFields count=\"").concat(columns.length === 0 ? 1 : columns.length, "\">\n        ").concat(columns.length === 0 ? '<field x="-2" />' : columns.map(columnIndex => "<field x=\"".concat(columnIndex, "\" />")).join('\n    '), "\n      </colFields>\n      <colItems count=\"1\">\n        <i t=\"grand\"><x /></i>\n      </colItems>\n      <dataFields count=\"").concat(values.length, "\">\n          ").concat(buildDataFields(cacheFields, values, metric), "\n      </dataFields>\n      <pivotTableStyleInfo\n        name=\"PivotStyleLight16\"\n        showRowHeaders=\"1\"\n        showColHeaders=\"1\"\n        showRowStripes=\"0\"\n        showColStripes=\"0\"\n        showLastColumn=\"1\"\n      />\n      <extLst>\n        <ext\n          uri=\"{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}\"\n          xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"\n        >\n          <x14:pivotTableDefinition\n            hideValuesRow=\"1\"\n            xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"\n          />\n        </ext>\n        <ext\n          uri=\"{747A6164-185A-40DC-8AA5-F01512510D54}\"\n          xmlns:xpdl=\"http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout\"\n        >\n          <xpdl:pivotTableDefinition16\n            EnabledSubtotalsDefault=\"0\"\n            SubtotalsOnTopDefault=\"0\"\n          />\n        </ext>\n      </extLst>\n    "));
    xmlStream.closeNode();
  }
  parseOpen(node) {
    // TK
  }
  parseText(text) {
    // TK
  }
  parseClose(name) {
    // TK
  }
  reconcile(model, options) {
    // TK
  }
}

// Helpers

function buildDataFields(cacheFields, values, metric) {
  let i = 0;
  let datafields = '';
  while (i < values.length) {
    const _metric = Array.isArray(metric) ? metric[i] : metric || 'sum';
    const nameKey = _metric.slice(0, 1).toUpperCase() + _metric.slice(1);
    datafields += "<dataField\n      name=\"".concat(nameKey, " of ").concat(cacheFields[values[i]].name, "\"\n      fld=\"").concat(values[i], "\"\n      subtotal=\"").concat(_metric, "\"\n      baseField=\"0\"\n      baseItem=\"0\"\n    />");
    i++;
  }
  return datafields;
}
function renderPivotFields(pivotTable) {
  /* eslint-disable no-nested-ternary */
  return pivotTable.cacheFields.map((cacheField, fieldIndex) => {
    const fieldType = pivotTable.rows.indexOf(fieldIndex) >= 0 ? 'row' : pivotTable.columns.indexOf(fieldIndex) >= 0 ? 'column' : pivotTable.values.indexOf(fieldIndex) >= 0 ? 'value' : null;
    return renderPivotField(fieldType, cacheField.sharedItems);
  }).join('');
}
function renderPivotField(fieldType, sharedItems) {
  // fieldType: 'row', 'column', 'value', null

  const defaultAttributes = 'compact="0" outline="0" showAll="0" defaultSubtotal="0"';
  if (fieldType === 'row' || fieldType === 'column') {
    const axis = fieldType === 'row' ? 'axisRow' : 'axisCol';
    return "\n      <pivotField axis=\"".concat(axis, "\" ").concat(defaultAttributes, ">\n        <items count=\"").concat(sharedItems.length + 1, "\">\n          ").concat(sharedItems.map((item, index) => "<item x=\"".concat(index, "\" />")).join('\n              '), "\n        </items>\n      </pivotField>\n    ");
  }
  return "\n    <pivotField\n      ".concat(fieldType === 'value' ? 'dataField="1"' : '', "\n      ").concat(defaultAttributes, "\n    />\n  ");
}
PivotTableXform.PIVOT_TABLE_ATTRIBUTES = {
  xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  'mc:Ignorable': 'xr',
  'xmlns:xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision'
};
module.exports = PivotTableXform;
//# sourceMappingURL=pivot-table-xform.js.map
