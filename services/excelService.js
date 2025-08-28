import fs from "fs";
import xlsx from "xlsx";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo de planilha não encontrado!");
  }

  const workbook = xlsx.readFile(excelPath);
  const firstSheetName = workbook.SheetNames[0];
  const csvData = xlsx.utils.sheet_to_csv(workbook.Sheets[firstSheetName]);
  const rows = csvData.split("\n");

  const extractedData = [];

  rows.forEach((row) => {
    const columns = row.split(",");

    const noteData = {
      dataTempo: columns[1],
      uf: columns[5],
      numDoc: columns[6],
      chave: columns[8],
      fornecedor: columns[12],
      autorizacao: columns[13],
      valorDoDoc: columns[16],
    };

    if (!noteData.chave || noteData.chave === "Chave de Acesso") return;

    extractedData.push(noteData);
  });

  return extractedData;
}

export { extractExcelData };
