// script.js
const API_BASE = ""; // Assumindo que o servidor serve o /public (mesmo host)
const form = document.getElementById("formInscricao");
const selectHorario = document.getElementById("horario");
const lista = document.getElementById("lista");
const mensagem = document.getElementById("mensagem");
const resumo = document.getElementById("resumo");
const btnRelatorio = document.getElementById("btnRelatorio");

async function fetchStatus() {
  const res = await fetch("/api/status");
  if (!res.ok) throw new Error("Falha ao obter status");
  const data = await res.json();
  return data;
}

async function populate() {
  try {
    const data = await fetchStatus();
    // preencher select com slots disponíveis
    selectHorario.innerHTML = "";
    const available = data.slots.filter(s => s.remaining > 0);
    if (available.length === 0) {
      const opt = document.createElement("option");
      opt.textContent = "Nenhum horário disponível";
      opt.disabled = true;
      selectHorario.appendChild(opt);
      selectHorario.disabled = true;
    } else {
      const placeholder = document.createElement("option");
      placeholder.textContent = "Escolha um horário";
      placeholder.value = "";
      placeholder.disabled = true;
      placeholder.selected = true;
      selectHorario.appendChild(placeholder);
      available.forEach(s => {
        const opt = document.createElement("option");
        opt.value = s.slot;
        opt.textContent = `${s.slot} — ${s.remaining} vagas`;
        selectHorario.appendChild(opt);
      });
      selectHorario.disabled = false;
    }

    // preencher lista visual com contagens por slot
    lista.innerHTML = "";
    data.slots.forEach(s => {
      const li = document.createElement("li");
      li.innerHTML = `<span>${s.slot}</span><strong>${s.remaining} vagas</strong>`;
      if (s.remaining === 0) li.querySelector("strong").textContent = "Cheio";
      lista.appendChild(li);
    });

    resumo.textContent = `Máximo por horário: ${data.maxPerSlot}. Horários: ${data.slots.length}.`;
  } catch (err) {
    console.error(err);
    mensagem.textContent = "Erro ao carregar horários.";
  }
}

form.addEventListener("submit", async (ev) => {
  ev.preventDefault();
  mensagem.textContent = "";
  const name = document.getElementById("nome").value.trim();
  const email = document.getElementById("email").value.trim();
  const slot = selectHorario.value;

  if (!name || !email || !slot) {
    mensagem.textContent = "Preencha todos os campos.";
    return;
  }

  try {
    const res = await fetch("/api/inscrever", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, email, slot })
    });
    const data = await res.json();
    if (!res.ok) {
      mensagem.textContent = data.error || "Erro ao inscrever";
      return;
    }
    mensagem.textContent = data.message || "Inscrição feita com sucesso!";
    form.reset();
    await populate();
  } catch (err) {
    console.error(err);
    mensagem.textContent = "Erro ao enviar inscrição.";
  }
});

btnRelatorio.addEventListener("click", () => {
  // abrir rota que fornece xlsx
  window.location.href = "/api/relatorio";
});

// Inicializa
populate();
