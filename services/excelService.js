import fs from "fs";
import xlsx from "xlsx";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo de planilha não encontrado!");
  }

  const workbook = xlsx.readFile(excelPath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  
  // Lê o Excel como uma matriz real de dados (evita quebra por vírgulas no texto)
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const extractedData = [];

  const limparCampo = (texto) => {
    if (texto === undefined || texto === null) return "";
    return String(texto).replace(/"/g, "").trim();
  };

  rows.forEach((columns) => {
    if (!columns || columns.length < 2) return;

    // Mapeamento baseado estritamente na ordem em que os dados apareceram na sua tabela
    const noteData = {
      dataTempo: limparCampo(columns[1]),   // Campo de data original (se houver)
      uf: limparCampo(columns[5]),          // UF original
      numDoc: limparCampo(columns[2]),      // Onde caiu o número do documento/valor
      chave: limparCampo(columns[3]),       // Onde caiu a chave de 44 dígitos
      fornecedor: limparCampo(columns[0]),  // Onde caiu a Razão Social/Fornecedor
      autorizacao: limparCampo(columns[4]), // Onde caiu o status de AUTORIZADA
      valorDoDoc: limparCampo(columns[6]),  // Onde caiu o valor final
    };

    // Ajuste dinâmico automático caso o SIGA inverta as posições das colunas:
    let chaveReal = "";
    columns.forEach((cell) => {
      const limpo = limparCampo(cell);
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;
      }
    });

    if (chaveReal) {
      noteData.chave = chaveReal;
    }

    // Ignora linhas de cabeçalho ou sem chave válida de NF-e
    if (!noteData.chave || noteData.chave.length < 20 || noteData.chave.toLowerCase().includes("chave")) {
      return;
    }

    extractedData.push(noteData);
  });

  return extractedData;
}

export { extractExcelData };
