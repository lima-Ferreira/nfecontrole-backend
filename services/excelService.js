import fs from "fs";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo não encontrado!");
  }

  const content = fs.readFileSync(excelPath, "utf-8");
  const rows = content.split(/\r?\n/);

  let extractedData = [];

  const limparCampo = (texto) => {
    if (!texto) return "";
    return texto.replace(/"/g, "").trim();
  };

  rows.forEach((row) => {
    if (!row || !row.trim()) return;

    if (row.includes("Razão Social") || row.includes("Chave NF-e")) return;

    const columns = row.includes(";") ? row.split(";") : row.split(",");

    let chaveReal = "";
    let fornecedorReal = "";
    let numeroDocReal = "";
    let dataReal = "";
    let valorReal = "";
    let ufReal = "CE";

    columns.forEach((cell, index) => {
      const limpo = limparCampo(cell);
      
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;

        // 1. Número do Documento arrancado direto da Chave de Acesso (Dígitos 25 a 34)
        const trechoNumero = chaveReal.substring(25, 34);
        numeroDocReal = String(parseInt(trechoNumero, 10));

        // 2. Captura o Fornecedor (Começo da linha)
        fornecedorReal = columns[1] ? limparCampo(columns[1]) : "";
        if (columns[2] && isNaN(limparCampo(columns[2])) && limparCampo(columns[2]).length > 2) {
          if (!limparCampo(columns[2]).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns[2]);
          }
        }

        // 3. Procura a UF (2 letras maiúsculas)
        for (let i = 1; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        // 4. Busca a Data de Emissão (DD/MM/AAAA)
        const campoData = columns.find(c => limparCampo(c).match(/^\d{2}\/\d{2}\/\d{4}$/));
        dataReal = campoData ? limparCampo(campoData) : "";

        // 5. Captura o Valor de forma posicional simples perto da chave
        if (columns[index - 1]) valorReal = limparCampo(columns[index - 1]);
        if ((!valorReal || valorReal.length > 12) && columns[index - 2]) {
          valorReal = limparCampo(columns[index - 2]);
        }
      }
    });

    if (chaveReal && chaveReal.length === 44) {
      extractedData.push({
        dataTempo: dataReal || "Ver no SIGA",
        uf: ufReal || "CE",
        numDoc: numeroDocReal || "NF-e",
        chave: chaveReal,
        fornecedor: fornecedorReal || "Fornecedor",
        autorizacao: "AUTORIZADA",
        valorDoDoc: valorReal || "0,00",
      });
    }
  });

  return extractedData;
}

export { extractExcelData };
