import fs from 'fs';
import { Node } from '@allanoricil/nrg-nodes';
import { isValidExtension } from '../utils';
import { template } from './template';

const ERROR_WRONG_DATA = 'Data must be an array of objects';
const ERROR_MISSING_TEMPLATE_FILEPATH = 'Template filepath is missing.';
const ERROR_MISSING_TEMPLATE_SHEET_NAME = 'Sheet is missing.';
const ERROR_MISSING_OUTPUT_FILEPATH = 'Output filepath is missing.';
const ERROR_WRONG_FILE_EXTENSION =
  'Filepath extension must be either .xlsx or .xls';
const ERROR_TEMPLATE_NOT_FOUND = 'Template not found';

export default class SpreadsheetTemplate extends Node {
  constructor(config) {
    super(config);

    this.templateFilepath = config.templateFilepath || '';
    this.templateSheetName = config.templateSheetName || '';
    this.outputFilepath = config.outputFilepath || '';
    this.templateFilepathType = config.templateFilepathType || 'str';
    this.templateSheetNameType = config.templateSheetNameType || 'str';
    this.outputFilepathType = config.outputFilepathType || 'str';
  }

  async onInput(msg, send, done) {
    try {
      this.status({});

      const data = msg.payload;
      if (
        !(
          Array.isArray(data) &&
          data.every((item) => typeof item === 'object' && item !== null)
        )
      ) {
        this.status({
          fill: 'red',
          shape: 'ring',
          text: ERROR_WRONG_DATA,
        });
        done(ERROR_WRONG_DATA);
        return;
      }

      const templateFilepath =
        msg.templateFilepath ||
        SpreadsheetTemplate.RED.util.evaluateNodeProperty(
          this.templateFilepath,
          this.templateFilepathType,
          this,
          msg,
        );

      const templateSheetName =
        msg.templateSheetName ||
        SpreadsheetTemplate.RED.util.evaluateNodeProperty(
          this.templateSheetName,
          this.templateSheetNameType,
          this,
          msg,
        );

      const outputFilepath =
        msg.outputFilepath ||
        SpreadsheetTemplate.RED.util.evaluateNodeProperty(
          this.outputFilepath,
          this.outputFilepathType,
          this,
          msg,
        );

      if (!templateFilepath) {
        this.status({
          fill: 'red',
          shape: 'ring',
          text: ERROR_MISSING_TEMPLATE_FILEPATH,
        });
        done(ERROR_MISSING_TEMPLATE_FILEPATH);
        return;
      }

      if (!isValidExtension(templateFilepath)) {
        this.status({
          fill: 'red',
          shape: 'ring',
          text: ERROR_WRONG_FILE_EXTENSION,
        });
        done(ERROR_WRONG_FILE_EXTENSION);
        return;
      }

      if (!templateSheetName) {
        this.status({
          fill: 'red',
          shape: 'ring',
          text: ERROR_MISSING_TEMPLATE_SHEET_NAME,
        });
        done(ERROR_MISSING_TEMPLATE_SHEET_NAME);
        return;
      }

      if (!outputFilepath) {
        this.status({
          fill: 'red',
          shape: 'ring',
          text: ERROR_MISSING_OUTPUT_FILEPATH,
        });
        done(ERROR_MISSING_OUTPUT_FILEPATH);
        return;
      }

      if (!isValidExtension(outputFilepath)) {
        this.status({
          fill: 'red',
          shape: 'ring',
          text: ERROR_WRONG_FILE_EXTENSION,
        });
        done(ERROR_WRONG_FILE_EXTENSION);
        return;
      }

      if (!fs.existsSync(templateFilepath)) {
        this.status({
          fill: 'red',
          shape: 'ring',
          text: ERROR_TEMPLATE_NOT_FOUND,
        });
        done(ERROR_TEMPLATE_NOT_FOUND);
        return;
      }

      const result = await template(
        templateFilepath,
        templateSheetName,
        data,
        outputFilepath,
      );

      if (result.error) {
        done(result.error);
        return;
      }

      this.log('spreadsheet created successfully');
      msg.payload = outputFilepath;
      this.status({
        fill: 'green',
        shape: 'dot',
        text: 'success',
      });
      send(msg);
      done();
    } catch (error) {
      done(error);
    } finally {
      setTimeout(() => {
        this.status({});
      }, 5 * 1000);
    }
  }
}
