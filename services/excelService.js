import fs from "fs";
import xlsx from "xlsx";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo de planilha não encontrado!");
  }

  const workbook = xlsx.readFile(excelPath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  
  // Transforma em matriz pura para ignorar as vírgulas dentro do nome do fornecedor
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const extractedData = [];

  const limparCampo = (texto) => {
    if (texto === undefined || texto === null) return "";
    return String(texto).replace(/"/g, "").trim();
  };

  rows.forEach((columns) => {
    if (!columns || columns.length < 5) return;

    // Procura dinamicamente a chave de 44 dígitos na linha caso ela mude de lugar
    let chaveReal = "";
    columns.forEach((cell) => {
      const limpo = limparCampo(cell);
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;
      }
    });

    // Se não encontrou uma chave de acesso válida de NF-e na linha, pula ela
    if (!chaveReal) return;

    // MAPEAMENTO CORRIGIDO PARA O SIGA (Índices reais da matriz)
    const noteData = {
      dataTempo: limparCampo(columns[0]),   // Coluna A (Geralmente Data ou Fornecedor no SIGA)
      uf: limparCampo(columns[1]),          // Coluna B (Geralmente UF ou Status)
      numDoc: limparCampo(columns[2]),      // Coluna C (Número do Documento)
      chave: chaveReal,                    // Chave de 44 dígitos encontrada dinamicamente
      fornecedor: limparCampo(columns[4]),  // Coluna E
      autorizacao: limparCampo(columns[5]), // Coluna F
      valorDoDoc: limparCampo(columns[6]),  // Coluna G
    };

    extractedData.push(noteData);
  });

  return extractedData;
}

export { extractExcelData };
