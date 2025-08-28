import express from "express";
import { fileUpload } from "./files-uploads/file.js";
import { extractExcelData } from "./services/excelService.js";
import { extractPdfChavesAsync } from "./services/pdfServices.js";
import ExcelJS from "exceljs";
import path from "path";
import cors from "cors";

const app = express();
app.use(
  cors({
    origin: "https://lima-ferreira.github.io", // seu frontend hospedado
    methods: ["GET", "POST"],
  })
);
const PORT = process.env.PORT || 7070;

// Rota de teste
app.get("/", (req, res) => {
  res.send("API rodando");
});

// API para upload e processamento de arquivos
app.post(
  "/api/upload",
  fileUpload.fields([
    { name: "excelFile", maxCount: 1 },
    { name: "pdfFile", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      // Validação: arquivos enviados
      const excelFile = req.files?.["excelFile"]?.[0];
      const pdfFile = req.files?.["pdfFile"]?.[0];

      if (!excelFile || !pdfFile) {
        return res.status(400).json({
          error: "É necessário enviar arquivos Excel e PDF.",
        });
      }

      const excelPath = excelFile.path;
      const pdfPath = pdfFile.path;

      // Processamento dos arquivos
      const arrExcel = extractExcelData(excelPath);
      const arrPdf = await extractPdfChavesAsync(pdfPath);

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

// API para download do Excel
app.get("/api/download-excel", async (req, res) => {
  try {
    const { chavesFaltando } = req.query;

    if (!chavesFaltando) return res.status(400).send("Nenhuma chave fornecida");

    const arr = JSON.parse(decodeURIComponent(chavesFaltando));

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

    arr.forEach((item) => {
      const row = worksheet.addRow(item);
      if (item.autorizacao?.toLowerCase() === "cancelado") {
        row.eachCell((cell) => {
          cell.font = { color: { argb: "FFFF0000" } }; // vermelho puro
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
    console.error(err);
    res.status(500).send("Erro ao gerar o Excel.");
  }
});

app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));
