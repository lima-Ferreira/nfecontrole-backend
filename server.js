import express from "express";
import cors from "cors";
import { fileUpload } from "./files-uploads/file.js";
import { extractExcelData } from "./services/excelService.js";
import { extractPdfChavesAsync } from "./services/pdfServices.js";
import ExcelJS from "exceljs";

const app = express();
app.use(cors());
app.use(express.json());

// Rota de upload e comparação
app.post(
  "/api/upload",
  fileUpload.fields([
    { name: "excelFile", maxCount: 1 },
    { name: "pdfFile", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const excelFile = req.files?.["excelFile"]?.[0];
      const pdfFile = req.files?.["pdfFile"]?.[0];

      if (!excelFile || !pdfFile) {
        return res.status(400).json({ error: "Envie Excel e PDF." });
      }

      const arrExcel = extractExcelData(excelFile.path);
      const arrPdf = await extractPdfChavesAsync(pdfFile.path);

      const chavesFaltando = arrExcel.filter(
        (item) => !arrPdf.includes(item.chave)
      );

      res.json({ chavesFaltando });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: "Erro ao processar arquivos." });
    }
  }
);

// Rota para download do Excel direto no navegador
app.post("/api/download-excel", async (req, res) => {
  try {
    const { chavesFaltando } = req.body;

    if (!Array.isArray(chavesFaltando) || chavesFaltando.length === 0) {
      return res.status(400).json({ error: "Nenhuma chave fornecida." });
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Chaves Faltando");

    worksheet.columns = [
      { header: "Data", key: "dataTempo", width: 15 },
      { header: "UF", key: "uf", width: 5 },
      { header: "Documento", key: "numDoc", width: 15 },
      { header: "Chave", key: "chave", width: 45 },
      { header: "Fornecedor", key: "fornecedor", width: 30 },
      { header: "Situação", key: "autorizacao", width: 15 },
      { header: "Valor", key: "valorDoDoc", width: 15 },
    ];

    worksheet.getRow(1).font = { bold: true };

    chavesFaltando.forEach((item) => {
      const row = worksheet.addRow(item);

      // Destacar canceladas
      if (item.autorizacao?.toLowerCase() === "cancelado") {
        row.eachCell((cell) => {
          cell.font = { color: { argb: "FFFF0000" } };
        });
      }
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="chaves_faltando.xlsx"'
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Erro ao gerar Excel:", err);
    res.status(500).json({ error: "Erro ao gerar Excel." });
  }
});

const PORT = process.env.PORT || 7070;
app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));
