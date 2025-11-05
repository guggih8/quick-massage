import express from "express";
import bodyParser from "body-parser";
import cors from "cors";
import ExcelJS from "exceljs";
import sqlite3 from "sqlite3";
import { open } from "sqlite";

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(bodyParser.json());
app.use(express.static("public"));

// ---- Conexão à base de dados SQLite ----
const db = new sqlite3.Database("./database.db", (err) => {
  if (err) console.error("Erro ao abrir BD:", err.message);
  else {
    db.run(`
      CREATE TABLE IF NOT EXISTS inscricoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        email TEXT NOT NULL,
        horario TEXT NOT NULL
      )
    `);
  }
});

// ---- Lista de horários disponíveis ----
const horarios = [
  "08:00", "08:30", "09:00", "09:30",
  "10:00", "10:30", "11:00", "11:30"
];

// ---- Obter horários disponíveis ----
app.get("/api/horarios", (req, res) => {
  db.all("SELECT horario, COUNT(*) as vagas FROM inscricoes GROUP BY horario", [], (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    const ocupacao = Object.fromEntries(rows.map(r => [r.horario, r.vagas]));
    const disponiveis = horarios.filter(h => !ocupacao[h] || ocupacao[h] < 4);
    res.json(disponiveis);
  });
});

// ---- Nova inscrição ----
app.post("/api/inscrever", (req, res) => {
  const { nome, email, horario } = req.body;
  if (!nome || !email || !horario) return res.status(400).json({ error: "Preencha todos os campos." });

  db.get("SELECT COUNT(*) as total FROM inscricoes WHERE horario = ?", [horario], (err, row) => {
    if (err) return res.status(500).json({ error: err.message });
    if (row.total >= 4) return res.status(400).json({ error: "Horário esgotado." });

    db.run("INSERT INTO inscricoes (nome, email, horario) VALUES (?, ?, ?)", [nome, email, horario], function (err2) {
      if (err2) return res.status(500).json({ error: err2.message });
      res.json({ success: true });
    });
  });
});

// ---- Exportar relatório Excel ----
app.get("/api/exportar", (req, res) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Inscrições");
  sheet.columns = [
    { header: "Nome", key: "nome", width: 25 },
    { header: "Email", key: "email", width: 30 },
    { header: "Horário", key: "horario", width: 10 }
  ];

  db.all("SELECT * FROM inscricoes ORDER BY horario", [], async (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    rows.forEach(i => sheet.addRow(i));

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=inscricoes_quick_massage.xlsx");
    await workbook.xlsx.write(res);
    res.end();
  });
});

app.listen(port, () => console.log(`Servidor a correr em http://localhost:${port}`));

