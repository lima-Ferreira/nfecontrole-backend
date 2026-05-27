import fs from "fs";
import xlsx from "xlsx";

function extractExcelData(excelPath) {
  if (!fs.existsSync(excelPath)) {
    throw new Error("Arquivo de planilha não encontrado!");
  }

  const workbook = xlsx.readFile(excelPath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  
  // Converte a planilha em JSON estruturado linha por linha
  // O parâmetro 'raw: false' força o formato de texto limpo
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });

  const extractedData = [];

  const limparCampo = (texto) => {
    if (texto === undefined || texto === null) return "";
    return String(texto).replace(/"/g, "").trim();
  };

  rows.forEach((columns) => {
    // Se a linha estiver vazia ou for o cabeçalho descritivo, pula
    if (!columns || columns.length === 0) return;

    let linhaTexto = String(columns[0]); // Pega o conteúdo bruto da Coluna A

    // Se for a linha dos títulos (CNPJ destinatário, Razão Social...), pula
    if (linhaTexto.includes("Razão Social") || linhaTexto.includes("Chave NF-e")) return;

    // Expressão regular mágica para dar split por vírgula APENAS fora das aspas
    const campos = linhaTexto.match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g) || linhaTexto.split(",");

    // Remove as aspas extras de cada pedaço encontrado
    const dadosLimpos = campos.map(c => limparCampo(c));

    // Com base no cabeçalho do SIGA visível no seu print:
    // [0] CNPJ | [1] Fornecedor | [2] UF | [3] NumDoc | [4] Data | [5] Indicadores | [6] Valor | [7] Chave
    const noteData = {
      dataTempo: dadosLimpos[4] || "",   // Data de Emissão
      uf: dadosLimpos[2] || "",          // UF
      numDoc: dadosLimpos[3] || "",      // Número da nota
      chave: dadosLimpos[7] || "",       // Chave NF-e (Último campo)
      fornecedor: dadosLimpos[1] || "",  // Razão Social destinatário
      autorizacao: dadosLimpos[5] || "AUTORIZADA", // Situação / Indicadores
      valorDoDoc: dadosLimpos[6] || "",  // Valor R$
    };

    // Validação estrita da chave de acesso (precisa ter 44 dígitos numéricos)
    if (!noteData.chave || noteData.chave.length !== 44 || isNaN(noteData.chave)) {
      // Procura em outros índices caso venha deslocado
      const chaveAlternativa = dadosLimpos.find(d => d.length === 44 && !isNaN(d));
      if (chaveAlternativa) {
        noteData.chave = chaveAlternativa;
      } else {
        return; // Pula se não achar chave válida
      }
    }

    extractedData.push(noteData);
  });

  return extractedData;
}

export { extractExcelData };
