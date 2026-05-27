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

        // 1. Número do Documento infalível (Extraído direto dos dígitos 25 a 34 da chave de acesso)
        // Remove os zeros à esquerda para ficar o número limpo da nota (ex: 180)
        const trechoNumero = chaveReal.substring(25, 34);
        numeroDocReal = String(parseInt(trechoNumero, 10));

        // 2. Captura o Fornecedor (Geralmente na coluna 1 ou 2)
        fornecedorReal = columns ? limparCampo(columns) : "";
        if (columns && isNaN(limparCampo(columns)) && limparCampo(columns).length > 2) {
          if (limparCampo(columns).length > 5 && !limparCampo(columns).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns);
          }
        }

        // 3. Procura a UF (2 letras maiúsculas)
        for (let i = 1; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        // 4. Busca a Data de Emissão (formato DD/MM/AAAA)
        const campoData = columns.find(c => limparCampo(c).match(/^\d{2}\/\d{2}\/\d{4}$/));
        dataReal = campoData ? limparCampo(campoData) : "";

        // 5. Caça o Valor correto R$ na linha (Pega o campo numérico que venha com ponto/vírgula perto da data/chave)
        for (let i = index - 1; i >= 2; i--) {
          const v = limparCampo(columns[i]);
          // O valor geralmente tem separador decimal e não é o lote gigante
          if (v && (v.includes(",") || v.includes(".")) && v.length > 2 && !v.includes("/") && v.length < 12) {
            valorReal = v;
            break;
          }
        }
        
        // Se não achou com ponto/vírgula, pega o campo imediatamente anterior à chave se ele for curto
        if (!valorReal && columns[index - 1]) {
          const vPrev = limparCampo(columns[index - 1]);
          if (vPrev.length < 10 && !isNaN(vPrev)) {
            valorReal = vPrev;
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

  // 🗓️ ORDENAÇÃO POR DATA DECRESCENTE PROTEGIDA
  extractedData.sort((a, b) => {
    try {
      const temDataA = a.dataTempo && a.dataTempo.includes("/");
      const temDataB = b.dataTempo && b.dataTempo.includes("/");

      if (!temDataA && !temDataB) return 0;
      if (!temDataA) return 1;
      if (!temDataB) return -1;

      const partesA = a.dataTempo.split("/");
      const partesB = b.dataTempo.split("/");

      const dataFormatadaA = new Date(`${partesA[2]}-${partesA[1]}-${partesA[0]}`);
      const dataFormatadaB = new Date(`${partesB[2]}-${partesB[1]}-${partesB[0]}`);

      if (isNaN(dataFormatadaA.getTime()) || isNaN(dataFormatadaB.getTime())) return 0;

      return dataFormatadaB - dataFormatadaA;
    } catch (e) {
      return 0;
    }
  });

  return extractedData;
}

export { extractExcelData };
