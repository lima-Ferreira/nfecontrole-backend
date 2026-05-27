import fs from "fs";
import xlsx from "xlsx";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo de planilha não encontrado!");
  }

  const workbook = xlsx.readFile(excelPath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  
  // Transforma em matriz pura para evitar problemas com vírgulas nos nomes dos fornecedores
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const extractedData = [];

  const limparCampo = (texto) => {
    if (texto === undefined || texto === null) return "";
    return String(texto).replace(/"/g, "").trim();
  };

  rows.forEach((columns) => {
    // Pula linhas vazias ou sem dados suficientes
    if (!columns || columns.length < 4) return;

    // Buscando as informações nas colunas corretas do novo formato do SIGA
    const noteData = {
      dataTempo: limparCampo(columns[0]),   // Primeira coluna onde está vindo a data ou fornecedor
      uf: limparCampo(columns[1]),          // Segunda coluna onde está vindo a UF ou Situação
      numDoc: limparCampo(columns[2]),      // Terceira coluna onde está vindo o Documento
      chave: limparCampo(columns[3]),       // Quarta coluna onde está vindo a Chave de 44 dígitos
      fornecedor: limparCampo(columns[4]),  // Quinta coluna
      autorizacao: limparCampo(columns[5]), // Sexta coluna
      valorDoDoc: limparCampo(columns[6]),  // Sétima coluna
    };

    // Caso a chave de 44 dígitos venha deslocada em outra coluna, fazemos uma varredura para garantir
    if (noteData.chave.length !== 44) {
      const chaveEncontrada = columns.find(cell => {
        const limpo = limparCampo(cell);
        return limpo.length === 44 && !isNaN(limpo);
      });
      if (chaveEncontrada) {
        noteData.chave = limparCampo(chaveEncontrada);
      }
    }

    // Regra rígida de validação: Só aceita se for uma Chave de Acesso válida de 44 números
    if (!noteData.chave || noteData.chave.length !== 44 || isNaN(noteData.chave)) {
      return;
    }

    extractedData.push(noteData);
  });

  return extractedData;
}

export { extractExcelData };
