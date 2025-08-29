import express from "express";
import cors from "cors";
import session from "express-session";
import ExcelJS from "exceljs";

const app = express();

// Middleware
app.use(cors()); // Permite requisições do frontend
app.use(express.json()); // Para ler JSON no body

// Sessão (se você ainda quiser armazenar dados temporários)
app.use(
  session({
    secret: "seu-segredo-aqui",
    resave: false,
    saveUninitialized: true,
  })
);

// Rota de teste da API
app.get("/", (req, res) => {
  res.send("API do NFE Controle rodando!");
});

// Rota para gerar Excel com POST
app.post("/api/download-excel", async (req, res) => {
  try {
    const chavesFaltando = req.body.chavesFaltando || [];

    if (!Array.isArray(chavesFaltando) || chavesFaltando.length === 0) {
      return res.status(400).json({ error: "Nenhuma chave fornecida" });
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
      if (item.autorizacao?.toLowerCase() === "cancelado") {
        row.eachCell((cell) => {
          cell.font = { color: { argb: "FFFF0000" } }; // vermelho
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

// Inicia o servidor
const PORT = process.env.PORT || 7070;
app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));
