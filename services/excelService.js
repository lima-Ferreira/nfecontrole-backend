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

        fornecedorReal = columns[1] ? limparCampo(columns[1]) : "";
        if (columns[2] && isNaN(limparCampo(columns[2])) && limparCampo(columns[2]).length > 2) {
          if (limparCampo(columns[2]).length > 5 && !limparCampo(columns[2]).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns[2]);
          }
        }

        for (let i = 1; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        for (let i = index - 1; i >= 0; i--) {
          const c = limparCampo(columns[i]);
          if (c && !isNaN(c) && c.length <= 9 && c.length > 0) {
            numeroDocReal = c;
            break;
          }
        }

        const campoData = columns.find(c => limparCampo(c).match(/^\d{2}\/\d{2}\/\d{4}$/));
        dataReal = campoData ? limparCampo(campoData) : "";

        for (let i = index - 1; i >= 2; i--) {
          const v = limparCampo(columns[i]);
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

  // 🗓️ ORDENAÇÃO POR DATA DECRESCENTE TOTALMENTE PROTEGIDA AGAINST CRASHES
  extractedData.sort((a, b) => {
    try {
      const temDataA = a.dataTempo && a.dataTempo.includes("/");
      const temDataB = b.dataTempo && b.dataTempo.includes("/");

      if (!temDataA && !temDataB) return 0;
      if (!temDataA) return 1;  // Joga registros sem data válida para o fim
      if (!temDataB) return -1; // Mantém registros com data no topo

      const partesA = a.dataTempo.split("/");
      const partesB = b.dataTempo.split("/");

      if (partesA.length !== 3 || partesB.length !== 3) return 0;

      const dataFormatadaA = new Date(`${partesA[2]}-${partesA[1]}-${partesA[0]}`);
      const dataFormatadaB = new Date(`${partesB[2]}-${partesB[1]}-${partesB[0]}`);

      // Verifica se a conversão gerou datas válidas antes de subtrair
      if (isNaN(dataFormatadaA.getTime()) || isNaN(dataFormatadaB.getTime())) return 0;

      return dataFormatadaB - dataFormatadaA; // Mais recentes primeiro
    } catch (sortError) {
      return 0; // Se falhar em alguma linha, não quebra a execução do servidor
    }
  });

  return extractedData;
}

export { extractExcelData };
