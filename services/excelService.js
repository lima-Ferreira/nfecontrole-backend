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
      
      // Identifica a Chave de Acesso
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;

        // Captura o fornecedor na primeira coluna preenchida
        fornecedorReal = columns ? limparCampo(columns) : "";
        if (columns && isNaN(limparCampo(columns)) && limparCampo(columns).length > 2) {
          if (limparCampo(columns).length > 5 && !limparCampo(columns).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns);
          }
        }

        // Procura a UF (2 letras maiúsculas)
        for (let i = 1; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        // Captura o Número do Documento correto (Procura um número menor isolado que não seja data)
        for (let i = index - 1; i >= 0; i--) {
          const c = limparCampo(columns[i]);
          if (c && !isNaN(c) && c.length <= 9 && c.length > 0) {
            numeroDocReal = c;
            break;
          }
        }

        // Busca a Data de Emissão (formato DD/MM/AAAA)
        const campoData = columns.find(c => limparCampo(c).match(/^\d{2}\/\d{2}\/\d{4}$/));
        dataReal = campoData ? limparCampo(campoData) : "";

        // Pega o Valor Real da Nota (Geralmente vem logo após os indicadores/data)
        for (let i = index - 1; i >= 2; i--) {
          const v = limparCampo(columns[i]);
          // Verifica se o campo tem cara de valor (ex: 150.00 ou 150,00 ou números com decimais)
          if (v && (v.includes(",") || v.includes(".") || !isNaN(v)) && v.length > 2 && v !== numeroDocReal && !v.includes("/")) {
            valorReal = v;
            break;
          }
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

  // 🗓️ ORDENAÇÃO POR DATA DECRESCENTE (Mais recente para a mais antiga)
  extractedData.sort((a, b) => {
    if (a.dataTempo.includes("/") && b.dataTempo.includes("/")) {
      const [diaA, mesA, anoA] = a.dataTempo.split("/");
      const [diaB, mesB, anoB] = b.dataTempo.split("/");
      
      const dataFormatadaA = new Date(`${anoA}-${mesA}-${diaA}`);
      const dataFormatadaB = new Date(`${anoB}-${mesB}-${diaB}`);
      
      return dataFormatadaB - dataFormatadaA; // Decrescente
    }
    return 0;
  });

  return extractedData;
}

export { extractExcelData };
