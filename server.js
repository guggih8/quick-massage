// server.js
import express from "express";
import bodyParser from "body-parser";
import cors from "cors";
import ExcelJS from "exceljs";
import Database from "better-sqlite3";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

/**
 * Configurações de horário
 * Sessões no dia 11/11/2025 entre 08:00 e 12:00,
 * duração 30 minutos => slots: 08:00, 08:30, ..., 11:30
 */
const SLOTS = [
  "08:00","08:30","09:00","09:30",
  "10:00","10:30","11:00","11:30"
];
const MAX_PER_SLOT = 4;
const EVENT_DATE = "2025-11-11"; // formato ISO (YYYY-MM-DD), para referência

app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, "public")));

// --- Inicializa DB (arquivo database.db) ---
const db = new Database(path.join(__dirname, "database.db"));

// Criar tabela se não existir
db.prepare(`
  CREATE TABLE IF NOT EXISTS registrations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT NOT NULL,
    slot TEXT NOT NULL,
    created_at TEXT NOT NULL
  )
`).run();

// Índice para consultas por slot/email
db.prepare(`CREATE INDEX IF NOT EXISTS idx_slot ON registrations(slot)`).run();
db.prepare(`CREATE INDEX IF NOT EXISTS idx_email ON registrations(email)`).run();

// --- Statements preparados ---
const countBySlotStmt = db.prepare("SELECT COUNT(*) AS cnt FROM registrations WHERE slot = ?");
const insertStmt = db.prepare("INSERT INTO registrations (name, email, slot, created_at) VALUES (?, ?, ?, ?)");
const selectAllStmt = db.prepare("SELECT * FROM registrations ORDER BY slot, created_at");

// Função utilitária para obter horários disponíveis
function getAvailableSlots() {
  const result = [];
  for (const s of SLOTS) {
    const row = countBySlotStmt.get(s);
    const count = row ? row.cnt : 0;
    if (count < MAX_PER_SLOT) result.push({ slot: s, remaining: MAX_PER_SLOT - count });
  }
  return result;
}

// --- Rotas API ---

// Retorna slots disponíveis (apenas horários ainda com vagas)
app.get("/api/horarios", (req, res) => {
  const avail = getAvailableSlots().map(x => ({ slot: x.slot, remaining: x.remaining }));
  res.json({ date: EVENT_DATE, slots: avail });
});

// Inscrição
app.post("/api/inscrever", (req, res) => {
  try {
    const { name, email, slot } = req.body || {};
    if (!name || !email || !slot) {
      return res.status(400).json({ error: "Campos obrigatórios: name, email, slot" });
    }

    if (!SLOTS.includes(slot)) {
      return res.status(400).json({ error: "Horário inválido" });
    }

    // Validação simples do email
    const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailPattern.test(email)) {
      return res.status(400).json({ error: "Email inválido" });
    }

    // Realizar operação em transacção para evitar race conditions
    const insertTx = db.transaction((nameInner, emailInner, slotInner) => {
      const row = countBySlotStmt.get(slotInner);
      const cnt = row ? row.cnt : 0;
      if (cnt >= MAX_PER_SLOT) {
        throw new Error("SLOT_FULL");
      }
      // Evita duplicação de email no mesmo slot
      const exists = db.prepare("SELECT 1 FROM registrations WHERE slot = ? AND email = ? LIMIT 1").get(slotInner, emailInner);
      if (exists) {
        throw new Error("ALREADY_REGISTERED");
      }
      const now = new Date().toISOString();
      insertStmt.run(nameInner, emailInner, slotInner, now);
      return true;
    });

    try {
      insertTx(name.trim(), email.trim().toLowerCase(), slot);
      return res.json({ ok: true, message: "Inscrição realizada com sucesso" });
    } catch (err) {
      if (err.message === "SLOT_FULL") return res.status(400).json({ error: "Horário já cheio" });
      if (err.message === "ALREADY_REGISTERED") return res.status(400).json({ error: "Email já inscrito neste horário" });
      console.error("Erro transacção:", err);
      return res.status(500).json({ error: "Erro ao inscrever" });
    }
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Erro servidor" });
  }
});

// Exportar relatório em formato Excel (.xlsx) - usamos exceljs para melhor compatibilidade
app.get("/api/relatorio", async (req, res) => {
  try {
    const rows = selectAllStmt.all(); // [{id,name,email,slot,created_at}, ...]

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Inscricoes");

    sheet.columns = [
      { header: "Horário", key: "slot", width: 12 },
      { header: "Nome", key: "name", width: 30 },
      { header: "Email", key: "email", width: 30 },
      { header: "Inscrito em (UTC)", key: "created_at", width: 25 }
    ];

    // Inserir linhas na mesma ordem dos slots
    for (const s of SLOTS) {
      const group = rows.filter(r => r.slot === s);
      if (group.length === 0) {
        // opcional: adicionar linha vazia para horário sem inscritos
        sheet.addRow({ slot: s, name: "", email: "", created_at: "" });
      } else {
        for (const r of group) {
          sheet.addRow({ slot: r.slot, name: r.name, email: r.email, created_at: r.created_at });
        }
      }
    }

    const filename = `relatorio_inscricoes_${EVENT_DATE}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Erro ao gerar relatório:", err);
    res.status(500).json({ error: "Falha ao gerar relatório" });
  }
});

// Rota que retorna contagens por slot (útil para frontend)
app.get("/api/status", (req, res) => {
  const status = SLOTS.map(s => {
    const row = countBySlotStmt.get(s);
    const cnt = row ? row.cnt : 0;
    return { slot: s, count: cnt, remaining: Math.max(0, MAX_PER_SLOT - cnt) };
  });
  res.json({ date: EVENT_DATE, slots: status, maxPerSlot: MAX_PER_SLOT });
});

// Serve index.html por padrão (já servis static em /public)
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// Start
app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});
