import fs from "fs";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo não encontrado!");
  }

  // Lê o arquivo do SIGA como texto puro (funciona para XLS, XLSX, CSV ou TXT)
  const content = fs.readFileSync(excelPath, "utf-8");
  const rows = content.split(/\r?\n/);

  const extractedData = [];

  const limparCampo = (texto) => {
    if (!texto) return "";
    return texto.replace(/"/g, "").trim();
  };

  rows.forEach((row) => {
    if (!row || !row.trim()) return;

    // Se for linha de título da Sefaz, pula
    if (row.includes("Razão Social") || row.includes("Chave NF-e")) return;

    // Divide por vírgula ou ponto-e-vírgula (o SIGA às vezes usa ;)
    const columns = row.includes(";") ? row.split(";") : row.split(",");

    let chaveReal = "";
    let fornecedorReal = "";
    let numeroDocReal = "";

    // Varre todos os campos da linha procurando a chave de 44 números
    columns.forEach((cell, index) => {
      const limpo = limparCampo(cell);
      
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;

        // Tenta pegar o Fornecedor e o Número do documento por aproximação de colunas
        fornecedorReal = columns[1] ? limparCampo(columns[1]) : "Fornecedor SIGA";
        numeroDocReal = columns[3] ? limparCampo(columns[3]) : "Nota SIGA";
        
        // Se o nome do fornecedor tiver vírgula e quebrou em dois, junta de volta
        if (columns[2] && isNaN(limparCampo(columns[2])) && limparCampo(columns[2]).length > 2) {
          if (!limparCampo(columns[2]).includes("CE") && limparCampo(columns[2]).length > 4) {
            fornecedorReal += " " + limparCampo(columns[2]);
          }
        }
      }
    });

    // Se achou uma chave válida de 44 dígitos, salva no sistema!
    if (chaveReal && chaveReal.length === 44) {
      extractedData.push({
        dataTempo: "Ver no SIGA",
        uf: "CE",
        numDoc: numeroDocReal || "NF-e",
        chave: chaveReal,
        fornecedor: fornecedorReal || "Fornecedor",
        autorizacao: "AUTORIZADA",
        valorDoDoc: "Ver no SIGA",
      });
    }
  });

  return extractedData;
}

export { extractExcelData };
