import fs from "fs";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo não encontrado!");
  }

  // Lê o arquivo do SIGA como texto puro
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

    // Divide a linha pelas vírgulas normais do CSV
    const columns = row.includes(";") ? row.split(";") : row.split(",");

    let chaveReal = "";
    let fornecedorReal = "";
    let numeroDocReal = "";
    let dataReal = "";
    let valorReal = "";
    let ufReal = "CE";

    // Varre todos os campos da linha procurando a chave de 44 números
    columns.forEach((cell, index) => {
      const limpo = limparCampo(cell);
      
      // Encontrou a chave de acesso!
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;

        // No padrão do SIGA, mapeamos os dados contando a partir do começo da linha:
        // O Fornecedor costuma ficar na coluna 1 ou 2
        fornecedorReal = columns[1] ? limparCampo(columns[1]) : "";
        
        // Se o nome do fornecedor tiver vírgula e quebrou em dois, junta de volta com a coluna 2
        if (columns[2] && isNaN(limparCampo(columns[2])) && limparCampo(columns[2]).length > 2) {
          if (limparCampo(columns[2]).length > 5 && !limparCampo(columns[2]).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns[2]);
          }
        }

        // Procura a UF (campo com 2 letras maiúsculas ex: CE, PE, SP) nas primeiras colunas
        for (let i = 2; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        // Pega o Número do Documento (geralmente vem algumas colunas antes da chave)
        if (columns[index - 4]) numeroDocReal = limparCampo(columns[index - 4]);

        // Busca a Data (formato DD/MM/AAAA ou similar) na linha
        const campoData = columns.find(c => limparCampo(c).match(/\d{2}\/\d{2}\/\d{4}/));
        dataReal = campoData ? limparCampo(campoData) : "Disponível no SIGA";

        // Busca o Valor R$ (procura um campo que tenha formato de dinheiro ou vírgula decimal perto do fim)
        if (columns[index - 1]) {
          const possivelValor = limparCampo(columns[index - 1]);
          if (possivelValor.includes(",") || !isNaN(possivelValor.replace(".", ""))) {
            valorReal = possivelValor;
          }
        }
        
        if (!valorReal && columns[index - 2]) {
          valorReal = limparCampo(columns[index - 2]);
        }
      }
    });

    // Se achou uma chave válida de 44 dígitos, salva no sistema com os dados reais extraídos!
    if (chaveReal && chaveReal.length === 44) {
      extractedData.push({
        dataTempo: dataReal || "Ver no SIGA",
        uf: ufReal || "CE",
        numDoc: numeroDocReal || "NF-e",
        chave: chaveReal,
        fornecedor: fornecedorReal || "Fornecedor",
        autorizacao: "AUTORIZADA",
        valorDoDoc: valorReal || "Ver no SIGA",
      });
    }
  });

  return extractedData;
}

export { extractExcelData };
import fs from "fs";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo não encontrado!");
  }

  // Lê o arquivo do SIGA como texto puro
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

    // Divide a linha pelas vírgulas normais do CSV
    const columns = row.includes(";") ? row.split(";") : row.split(",");

    let chaveReal = "";
    let fornecedorReal = "";
    let numeroDocReal = "";
    let dataReal = "";
    let valorReal = "";
    let ufReal = "CE";

    // Varre todos os campos da linha procurando a chave de 44 números
    columns.forEach((cell, index) => {
      const limpo = limparCampo(cell);
      
      // Encontrou a chave de acesso!
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;

        // No padrão do SIGA, mapeamos os dados contando a partir do começo da linha:
        // O Fornecedor costuma ficar na coluna 1 ou 2
        fornecedorReal = columns[1] ? limparCampo(columns[1]) : "";
        
        // Se o nome do fornecedor tiver vírgula e quebrou em dois, junta de volta com a coluna 2
        if (columns[2] && isNaN(limparCampo(columns[2])) && limparCampo(columns[2]).length > 2) {
          if (limparCampo(columns[2]).length > 5 && !limparCampo(columns[2]).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns[2]);
          }
        }

        // Procura a UF (campo com 2 letras maiúsculas ex: CE, PE, SP) nas primeiras colunas
        for (let i = 2; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        // Pega o Número do Documento (geralmente vem algumas colunas antes da chave)
        if (columns[index - 4]) numeroDocReal = limparCampo(columns[index - 4]);

        // Busca a Data (formato DD/MM/AAAA ou similar) na linha
        const campoData = columns.find(c => limparCampo(c).match(/\d{2}\/\d{2}\/\d{4}/));
        dataReal = campoData ? limparCampo(campoData) : "Disponível no SIGA";

        // Busca o Valor R$ (procura um campo que tenha formato de dinheiro ou vírgula decimal perto do fim)
        if (columns[index - 1]) {
          const possivelValor = limparCampo(columns[index - 1]);
          if (possivelValor.includes(",") || !isNaN(possivelValor.replace(".", ""))) {
            valorReal = possivelValor;
          }
        }
        
        if (!valorReal && columns[index - 2]) {
          valorReal = limparCampo(columns[index - 2]);
        }
      }
    });

    // Se achou uma chave válida de 44 dígitos, salva no sistema com os dados reais extraídos!
    if (chaveReal && chaveReal.length === 44) {
      extractedData.push({
        dataTempo: dataReal || "Ver no SIGA",
        uf: ufReal || "CE",
        numDoc: numeroDocReal || "NF-e",
        chave: chaveReal,
        fornecedor: fornecedorReal || "Fornecedor",
        autorizacao: "AUTORIZADA",
        valorDoDoc: valorReal || "Ver no SIGA",
      });
    }
  });

  return extractedData;
}

export { extractExcelData };
import fs from "fs";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo não encontrado!");
  }

  // Lê o arquivo do SIGA como texto puro
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

    // Divide a linha pelas vírgulas normais do CSV
    const columns = row.includes(";") ? row.split(";") : row.split(",");

    let chaveReal = "";
    let fornecedorReal = "";
    let numeroDocReal = "";
    let dataReal = "";
    let valorReal = "";
    let ufReal = "CE";

    // Varre todos os campos da linha procurando a chave de 44 números
    columns.forEach((cell, index) => {
      const limpo = limparCampo(cell);
      
      // Encontrou a chave de acesso!
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;

        // No padrão do SIGA, mapeamos os dados contando a partir do começo da linha:
        // O Fornecedor costuma ficar na coluna 1 ou 2
        fornecedorReal = columns[1] ? limparCampo(columns[1]) : "";
        
        // Se o nome do fornecedor tiver vírgula e quebrou em dois, junta de volta com a coluna 2
        if (columns[2] && isNaN(limparCampo(columns[2])) && limparCampo(columns[2]).length > 2) {
          if (limparCampo(columns[2]).length > 5 && !limparCampo(columns[2]).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns[2]);
          }
        }

        // Procura a UF (campo com 2 letras maiúsculas ex: CE, PE, SP) nas primeiras colunas
        for (let i = 2; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        // Pega o Número do Documento (geralmente vem algumas colunas antes da chave)
        if (columns[index - 4]) numeroDocReal = limparCampo(columns[index - 4]);

        // Busca a Data (formato DD/MM/AAAA ou similar) na linha
        const campoData = columns.find(c => limparCampo(c).match(/\d{2}\/\d{2}\/\d{4}/));
        dataReal = campoData ? limparCampo(campoData) : "Disponível no SIGA";

        // Busca o Valor R$ (procura um campo que tenha formato de dinheiro ou vírgula decimal perto do fim)
        if (columns[index - 1]) {
          const possivelValor = limparCampo(columns[index - 1]);
          if (possivelValor.includes(",") || !isNaN(possivelValor.replace(".", ""))) {
            valorReal = possivelValor;
          }
        }
        
        if (!valorReal && columns[index - 2]) {
          valorReal = limparCampo(columns[index - 2]);
        }
      }
    });

    // Se achou uma chave válida de 44 dígitos, salva no sistema com os dados reais extraídos!
    if (chaveReal && chaveReal.length === 44) {
      extractedData.push({
        dataTempo: dataReal || "Ver no SIGA",
        uf: ufReal || "CE",
        numDoc: numeroDocReal || "NF-e",
        chave: chaveReal,
        fornecedor: fornecedorReal || "Fornecedor",
        autorizacao: "AUTORIZADA",
        valorDoDoc: valorReal || "Ver no SIGA",
      });
    }
  });

  return extractedData;
}

export { extractExcelData };
import fs from "fs";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo não encontrado!");
  }

  // Lê o arquivo do SIGA como texto puro
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

    // Divide a linha pelas vírgulas normais do CSV
    const columns = row.includes(";") ? row.split(";") : row.split(",");

    let chaveReal = "";
    let fornecedorReal = "";
    let numeroDocReal = "";
    let dataReal = "";
    let valorReal = "";
    let ufReal = "CE";

    // Varre todos os campos da linha procurando a chave de 44 números
    columns.forEach((cell, index) => {
      const limpo = limparCampo(cell);
      
      // Encontrou a chave de acesso!
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;

        // No padrão do SIGA, mapeamos os dados contando a partir do começo da linha:
        // O Fornecedor costuma ficar na coluna 1 ou 2
        fornecedorReal = columns[1] ? limparCampo(columns[1]) : "";
        
        // Se o nome do fornecedor tiver vírgula e quebrou em dois, junta de volta com a coluna 2
        if (columns[2] && isNaN(limparCampo(columns[2])) && limparCampo(columns[2]).length > 2) {
          if (limparCampo(columns[2]).length > 5 && !limparCampo(columns[2]).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns[2]);
          }
        }

        // Procura a UF (campo com 2 letras maiúsculas ex: CE, PE, SP) nas primeiras colunas
        for (let i = 2; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        // Pega o Número do Documento (geralmente vem algumas colunas antes da chave)
        if (columns[index - 4]) numeroDocReal = limparCampo(columns[index - 4]);

        // Busca a Data (formato DD/MM/AAAA ou similar) na linha
        const campoData = columns.find(c => limparCampo(c).match(/\d{2}\/\d{2}\/\d{4}/));
        dataReal = campoData ? limparCampo(campoData) : "Disponível no SIGA";

        // Busca o Valor R$ (procura um campo que tenha formato de dinheiro ou vírgula decimal perto do fim)
        if (columns[index - 1]) {
          const possivelValor = limparCampo(columns[index - 1]);
          if (possivelValor.includes(",") || !isNaN(possivelValor.replace(".", ""))) {
            valorReal = possivelValor;
          }
        }
        
        if (!valorReal && columns[index - 2]) {
          valorReal = limparCampo(columns[index - 2]);
        }
      }
    });

    // Se achou uma chave válida de 44 dígitos, salva no sistema com os dados reais extraídos!
    if (chaveReal && chaveReal.length === 44) {
      extractedData.push({
        dataTempo: dataReal || "Ver no SIGA",
        uf: ufReal || "CE",
        numDoc: numeroDocReal || "NF-e",
        chave: chaveReal,
        fornecedor: fornecedorReal || "Fornecedor",
        autorizacao: "AUTORIZADA",
        valorDoDoc: valorReal || "Ver no SIGA",
      });
    }
  });

  return extractedData;
}

export { extractExcelData };
import fs from "fs";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo não encontrado!");
  }

  // Lê o arquivo do SIGA como texto puro
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

    // Divide a linha pelas vírgulas normais do CSV
    const columns = row.includes(";") ? row.split(";") : row.split(",");

    let chaveReal = "";
    let fornecedorReal = "";
    let numeroDocReal = "";
    let dataReal = "";
    let valorReal = "";
    let ufReal = "CE";

    // Varre todos os campos da linha procurando a chave de 44 números
    columns.forEach((cell, index) => {
      const limpo = limparCampo(cell);
      
      // Encontrou a chave de acesso!
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;

        // No padrão do SIGA, mapeamos os dados contando a partir do começo da linha:
        // O Fornecedor costuma ficar na coluna 1 ou 2
        fornecedorReal = columns[1] ? limparCampo(columns[1]) : "";
        
        // Se o nome do fornecedor tiver vírgula e quebrou em dois, junta de volta com a coluna 2
        if (columns[2] && isNaN(limparCampo(columns[2])) && limparCampo(columns[2]).length > 2) {
          if (limparCampo(columns[2]).length > 5 && !limparCampo(columns[2]).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns[2]);
          }
        }

        // Procura a UF (campo com 2 letras maiúsculas ex: CE, PE, SP) nas primeiras colunas
        for (let i = 2; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        // Pega o Número do Documento (geralmente vem algumas colunas antes da chave)
        if (columns[index - 4]) numeroDocReal = limparCampo(columns[index - 4]);

        // Busca a Data (formato DD/MM/AAAA ou similar) na linha
        const campoData = columns.find(c => limparCampo(c).match(/\d{2}\/\d{2}\/\d{4}/));
        dataReal = campoData ? limparCampo(campoData) : "Disponível no SIGA";

        // Busca o Valor R$ (procura um campo que tenha formato de dinheiro ou vírgula decimal perto do fim)
        if (columns[index - 1]) {
          const possivelValor = limparCampo(columns[index - 1]);
          if (possivelValor.includes(",") || !isNaN(possivelValor.replace(".", ""))) {
            valorReal = possivelValor;
          }
        }
        
        if (!valorReal && columns[index - 2]) {
          valorReal = limparCampo(columns[index - 2]);
        }
      }
    });

    // Se achou uma chave válida de 44 dígitos, salva no sistema com os dados reais extraídos!
    if (chaveReal && chaveReal.length === 44) {
      extractedData.push({
        dataTempo: dataReal || "Ver no SIGA",
        uf: ufReal || "CE",
        numDoc: numeroDocReal || "NF-e",
        chave: chaveReal,
        fornecedor: fornecedorReal || "Fornecedor",
        autorizacao: "AUTORIZADA",
        valorDoDoc: valorReal || "Ver no SIGA",
      });
    }
  });

  return extractedData;
}

export { extractExcelData };
import fs from "fs";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo não encontrado!");
  }

  // Lê o arquivo do SIGA como texto puro
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

    // Divide a linha pelas vírgulas normais do CSV
    const columns = row.includes(";") ? row.split(";") : row.split(",");

    let chaveReal = "";
    let fornecedorReal = "";
    let numeroDocReal = "";
    let dataReal = "";
    let valorReal = "";
    let ufReal = "CE";

    // Varre todos os campos da linha procurando a chave de 44 números
    columns.forEach((cell, index) => {
      const limpo = limparCampo(cell);
      
      // Encontrou a chave de acesso!
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;

        // No padrão do SIGA, mapeamos os dados contando a partir do começo da linha:
        // O Fornecedor costuma ficar na coluna 1 ou 2
        fornecedorReal = columns[1] ? limparCampo(columns[1]) : "";
        
        // Se o nome do fornecedor tiver vírgula e quebrou em dois, junta de volta com a coluna 2
        if (columns[2] && isNaN(limparCampo(columns[2])) && limparCampo(columns[2]).length > 2) {
          if (limparCampo(columns[2]).length > 5 && !limparCampo(columns[2]).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns[2]);
          }
        }

        // Procura a UF (campo com 2 letras maiúsculas ex: CE, PE, SP) nas primeiras colunas
        for (let i = 2; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        // Pega o Número do Documento (geralmente vem algumas colunas antes da chave)
        if (columns[index - 4]) numeroDocReal = limparCampo(columns[index - 4]);

        // Busca a Data (formato DD/MM/AAAA ou similar) na linha
        const campoData = columns.find(c => limparCampo(c).match(/\d{2}\/\d{2}\/\d{4}/));
        dataReal = campoData ? limparCampo(campoData) : "Disponível no SIGA";

        // Busca o Valor R$ (procura um campo que tenha formato de dinheiro ou vírgula decimal perto do fim)
        if (columns[index - 1]) {
          const possivelValor = limparCampo(columns[index - 1]);
          if (possivelValor.includes(",") || !isNaN(possivelValor.replace(".", ""))) {
            valorReal = possivelValor;
          }
        }
        
        if (!valorReal && columns[index - 2]) {
          valorReal = limparCampo(columns[index - 2]);
        }
      }
    });

    // Se achou uma chave válida de 44 dígitos, salva no sistema com os dados reais extraídos!
    if (chaveReal && chaveReal.length === 44) {
      extractedData.push({
        dataTempo: dataReal || "Ver no SIGA",
        uf: ufReal || "CE",
        numDoc: numeroDocReal || "NF-e",
        chave: chaveReal,
        fornecedor: fornecedorReal || "Fornecedor",
        autorizacao: "AUTORIZADA",
        valorDoDoc: valorReal || "Ver no SIGA",
      });
    }
  });

  return extractedData;
}

export { extractExcelData };
import fs from "fs";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo não encontrado!");
  }

  // Lê o arquivo do SIGA como texto puro
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

    // Divide a linha pelas vírgulas normais do CSV
    const columns = row.includes(";") ? row.split(";") : row.split(",");

    let chaveReal = "";
    let fornecedorReal = "";
    let numeroDocReal = "";
    let dataReal = "";
    let valorReal = "";
    let ufReal = "CE";

    // Varre todos os campos da linha procurando a chave de 44 números
    columns.forEach((cell, index) => {
      const limpo = limparCampo(cell);
      
      // Encontrou a chave de acesso!
      if (limpo.length === 44 && !isNaN(limpo)) {
        chaveReal = limpo;

        // No padrão do SIGA, mapeamos os dados contando a partir do começo da linha:
        // O Fornecedor costuma ficar na coluna 1 ou 2
        fornecedorReal = columns[1] ? limparCampo(columns[1]) : "";
        
        // Se o nome do fornecedor tiver vírgula e quebrou em dois, junta de volta com a coluna 2
        if (columns[2] && isNaN(limparCampo(columns[2])) && limparCampo(columns[2]).length > 2) {
          if (limparCampo(columns[2]).length > 5 && !limparCampo(columns[2]).match(/^[A-Z]{2}$/)) {
            fornecedorReal += " " + limparCampo(columns[2]);
          }
        }

        // Procura a UF (campo com 2 letras maiúsculas ex: CE, PE, SP) nas primeiras colunas
        for (let i = 2; i < 5; i++) {
          if (columns[i] && limparCampo(columns[i]).match(/^[A-Z]{2}$/)) {
            ufReal = limparCampo(columns[i]);
            break;
          }
        }

        // Pega o Número do Documento (geralmente vem algumas colunas antes da chave)
        if (columns[index - 4]) numeroDocReal = limparCampo(columns[index - 4]);

        // Busca a Data (formato DD/MM/AAAA ou similar) na linha
        const campoData = columns.find(c => limparCampo(c).match(/\d{2}\/\d{2}\/\d{4}/));
        dataReal = campoData ? limparCampo(campoData) : "Disponível no SIGA";

        // Busca o Valor R$ (procura um campo que tenha formato de dinheiro ou vírgula decimal perto do fim)
        if (columns[index - 1]) {
          const possivelValor = limparCampo(columns[index - 1]);
          if (possivelValor.includes(",") || !isNaN(possivelValor.replace(".", ""))) {
            valorReal = possivelValor;
          }
        }
        
        if (!valorReal && columns[index - 2]) {
          valorReal = limparCampo(columns[index - 2]);
        }
      }
    });

    // Se achou uma chave válida de 44 dígitos, salva no sistema com os dados reais extraídos!
    if (chaveReal && chaveReal.length === 44) {
      extractedData.push({
        dataTempo: dataReal || "Ver no SIGA",
        uf: ufReal || "CE",
        numDoc: numeroDocReal || "NF-e",
        chave: chaveReal,
        fornecedor: fornecedorReal || "Fornecedor",
        autorizacao: "AUTORIZADA",
        valorDoDoc: valorReal || "Ver no SIGA",
      });
    }
  });

  return extractedData;
}

export { extractExcelData };
