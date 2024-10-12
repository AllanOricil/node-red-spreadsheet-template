import * as ExcelJS from 'exceljs';

async function template(
  templateFilepath,
  templateSheetName,
  data,
  outputFilepath,
) {
  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.readFile(templateFilepath);
  const worksheet = workbook.getWorksheet(templateSheetName);

  if (!worksheet) {
    return {
      error: `Could not find ${templateSheetName} in the template`,
    };
  }

  const headers = [];
  worksheet.getRow(1).eachCell((cell) => {
    headers.push(cell.value);
  });

  const firstDataRow = 2;
  const templateRow = worksheet.getRow(firstDataRow);
  data.forEach((row, rowIndex) => {
    const newRow = worksheet.getRow(firstDataRow + rowIndex);
    headers.forEach((header, colIndex) => {
      const templateCell = templateRow.getCell(colIndex + 1);
      const cell = newRow.getCell(colIndex + 1);
      cell.alignment = templateCell.alignment;
      cell.border = templateCell.border;
      cell.dataValidation = templateCell.dataValidation;
      cell.effectiveType = templateCell.effectiveType;
      cell.fill = templateCell.fill;
      cell.font = templateCell.font;
      cell.formula = templateCell.formula;
      cell.formulaType = templateCell.formulaType;
      cell.hyperlink = templateCell.hyperlink;
      cell.name = templateCell.name;
      cell.names = templateCell.names;
      cell.numFmt = templateCell.numFmt;
      cell.protection = templateCell.protection;
      cell.type = templateCell.type;

      cell.value = templateCell.hyperlink
        ? {
            text: row[header],
            // NOTE: email is usually the only url type that isn't provided with its protocol
            hyperlink: /^[\w-.]+@([\w-]+\.)+[\w-]{2,4}$/.test(row[header])
              ? `mailto:${row[header]}`
              : row[header],
          }
        : row[header];
    });
    newRow.commit();
  });

  await workbook.xlsx.writeFile(outputFilepath);

  return { outputFilepath };
}

export { template };
