import express from "express";
import session from "express-session";
import { fileUpload } from "./files-uploads/file.js";
import { extractExcelData } from "./services/excelService.js";
import { extractPdfChavesAsync } from "./services/pdfServices.js";
import ExcelJS from "exceljs";
import path from "path";

const app = express();

// ================== MIDDLEWARES ================== //
app.use(express.json()); // para receber JSON no POST
app.use(express.urlencoded({ extended: true }));
app.use(express.static("public"));

// Configura sessão
app.use(
  session({
    secret: "seu-segredo-aqui",
    resave: false,
    saveUninitialized: true,
  })
);

// ================== ROTAS ================== //

// Página inicial
app.get("/", (req, res) => {
  const chavesFaltando = req.session.chavesFaltando || [];
  res.render("home", { chavesFaltando });
});

// Página de resultados
app.get("/nfe", (req, res) => {
  const chavesFaltando = req.session.chavesFaltando || [];
  res.render("nfe.ejs", { chavesFaltando });
});

// Upload de arquivos
app.post(
  "/upload",
  fileUpload.fields([
    { name: "excelFile", maxCount: 1 },
    { name: "pdfFile", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const excelPath = req.files["excelFile"][0].path;
      const pdfPath = req.files["pdfFile"][0].path;

      const arrExcel = extractExcelData(excelPath);
      const arrPdf = await extractPdfChavesAsync(pdfPath);

      // Compara e guarda chaves faltantes na sessão
      const chavesFaltando = arrExcel.filter(
        (item) => !arrPdf.includes(item.chave)
      );

      req.session.chavesFaltando = chavesFaltando;

      res.redirect("/nfe");
    } catch (err) {
      console.error(err);
      res.status(500).send("Erro ao processar arquivos.");
    }
  }
);

// ================== ROTA PARA DOWNLOAD DO EXCEL VIA POST ================== //
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
// ================== INICIA O SERVIDOR ================== //
const PORT = process.env.PORT || 7070;
app.listen(PORT, () =>
  console.log(`Servidor rodando na porta ${PORT}: http://localhost:${PORT}`)
);
