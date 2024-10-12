import { validateFilepathInput } from './utils';

export default {
  category: 'spreadsheet',
  color: '#3cb277',
  icon: 'icon.png',
  inputs: 1,
  outputs: 1,
  defaults: {
    name: {
      value: '',
    },
    templateFilepath: {
      value: 'template.xlsx',
      required: true,
      validate: (value) =>
        validateFilepathInput(value, '#node-input-templateFilepathType'),
    },
    templateSheetName: {
      value: 'sheet-1',
      required: true,
    },
    outputFilepath: {
      value: 'output.xlsx',
      required: true,
      validate: (value) =>
        validateFilepathInput(value, '#node-input-outputFilepathType'),
    },
    templateFilepathType: {
      value: 'str',
    },
    templateSheetNameType: {
      value: 'str',
    },
    outputFilepathType: {
      value: 'str',
    },
  },
  paletteLabel: function () {
    return this.name || this._('spreadsheet-template.name') || 'template';
  },
  label: function () {
    return this.name || this._('spreadsheet-template.name') || 'template';
  },
  inputLabels: function () {
    return this._('spreadsheet-template.input.name') || 'data';
  },
  outputLabels: function () {
    return `${this._('spreadsheet-template.outputs.outputFilepath.name') || 'output filepath'}`;
  },
  oneditprepare: function () {
    $('#node-input-templateFilepath').typedInput({
      type: 'msg',
      types: ['msg', 'str'],
      typeField: '#node-input-templateFilepathType',
    });
    $('#node-input-templateSheetName').typedInput({
      type: 'msg',
      types: ['msg', 'str'],
      typeField: '#node-input-templateSheetNameType',
    });
    $('#node-input-outputFilepath').typedInput({
      type: 'msg',
      types: ['msg', 'str'],
      typeField: '#node-input-outputFilepathType',
    });
  },
};
