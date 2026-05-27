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

  // Função para remover as aspas duplicadas que o SIGA gera
  const limparCampo = (texto) => {
    if (!texto) return "";
    return texto.replace(/"/g, "").trim();
  };

  rows.forEach((row) => {
    if (!row.trim()) return;

    const columns = row.split(",");

    // Reorganização baseada no novo layout do SIGA que vimos no seu print
    const noteData = {
      dataTempo: limparCampo(columns[1]),   // Ajuste se a data mudou de lugar na planilha
      uf: limparCampo(columns[5]),          
      numDoc: limparCampo(columns[6]),      
      chave: limparCampo(columns[1]),       // No print, a chave veio parar onde era a coluna do documento
      fornecedor: limparCampo(columns[0]),  // No print, o fornecedor veio parar na primeira coluna (onde era a data)
      autorizacao: limparCampo(columns[2]), // No print, a situação veio parar onde era a UF
      valorDoDoc: limparCampo(columns[7]),  // Ajuste estimado para o valor
    };

    // Ajuste inteligente: Se o fornecedor veio na coluna 0 e a chave na coluna 1 (como no print quebrado):
    if (columns[1] && columns[1].replace(/"/g, "").trim().length === 44) {
      noteData.chave = limparCampo(columns[1]);
      noteData.fornecedor = limparCampo(columns[0]);
      noteData.autorizacao = limparCampo(columns[2]);
      // Deixamos a data fixa ou tentamos buscar de outra coluna se o SIGA mudou ela de lugar
      noteData.dataTempo = "Ver no SIGA"; 
    }

    // Ignora linhas de cabeçalho ou sem chave válida
    if (!noteData.chave || noteData.chave.toLowerCase().includes("chave") || noteData.chave.length < 20) {
      return;
    }

    extractedData.push(noteData);
  });

  return extractedData;
}

export { extractExcelData };
