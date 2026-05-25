const apiKeyInput = document.getElementById("apiKey");
const toggleKeyBtn = document.getElementById("toggleKey");
const descricaoInput = document.getElementById("descricao");
const processarBtn = document.getElementById("processarBtn");
const salvarBtn = document.getElementById("salvarBtn");
const importarBtn = document.getElementById("importarBtn");
const arquivoInput = document.getElementById("arquivoInput");
const statusMessage = document.getElementById("statusMessage");
const contadorEl = document.getElementById("contador");
const abaLabel = document.getElementById("abaLabel");

const campos = {
  nome: document.getElementById("nome"),
  sobrenome: document.getElementById("sobrenome"),
  cidade: document.getElementById("cidade"),
  telefone: document.getElementById("telefone"),
  fonte: document.getElementById("fonte"),
  status: document.getElementById("status"),
  observacoes: document.getElementById("observacoes"),
};

function setStatus(message) {
  statusMessage.textContent = message;
}

function formatContador(total) {
  const texto = total === 1 ? "registro salvo hoje" : "registros salvos hoje";
  return `${total} ${texto}`;
}

async function atualizarContador() {
  try {
    const response = await fetch("/api/contador");
    if (!response.ok) {
      return;
    }
    const data = await response.json();
    contadorEl.textContent = formatContador(data.total || 0);
  } catch (error) {
    console.warn(error);
  }
}

function abaPorStatus(status) {
  const valor = (status || "").toLowerCase();
  if (valor.includes("interesse") || valor.includes("recusou")) {
    return 'Sem Interesse';
  }
  return 'Falhas e Sem Contato';
}

function atualizarAba() {
  const aba = abaPorStatus(campos.status.value);
  abaLabel.textContent = `→ será salvo em: "${aba}"`;
}

function preencherCampos(dados) {
  campos.nome.value = dados.nome || "";
  campos.sobrenome.value = dados.sobrenome || "";
  campos.cidade.value = dados.cidade || "";
  campos.telefone.value = dados.telefone || "";
  campos.fonte.value = dados.fonte || window.APP_CONFIG.fontes[0];
  campos.status.value = dados.status || window.APP_CONFIG.statuses[0];
  campos.observacoes.value = dados.observacoes || "";
  atualizarAba();
}

function limparCampos() {
  descricaoInput.value = "";
  preencherCampos({});
}

async function processar() {
  const texto = descricaoInput.value.trim();
  if (!texto) {
    setStatus("Descreva o contato antes de processar.");
    return;
  }

  processarBtn.disabled = true;
  setStatus("⏳ Chamando a IA — aguarde...");

  try {
    const response = await fetch("/api/processar", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        texto,
        api_key: apiKeyInput.value.trim(),
      }),
    });
    const data = await response.json();
    if (!response.ok) {
      setStatus(`✗ ${data.error || "Erro ao chamar a IA."}`);
      return;
    }
    preencherCampos(data.dados || {});
    setStatus("✓ Dados extraídos. Revise os campos e salve.");
  } catch (error) {
    console.error(error);
    setStatus("✗ Erro ao chamar a IA.");
  } finally {
    processarBtn.disabled = false;
  }
}

async function salvar() {
  const payload = {
    nome: campos.nome.value.trim(),
    sobrenome: campos.sobrenome.value.trim(),
    cidade: campos.cidade.value.trim(),
    telefone: campos.telefone.value.trim(),
    fonte: campos.fonte.value.trim(),
    status: campos.status.value.trim(),
    observacoes: campos.observacoes.value.trim(),
  };

  if (!payload.nome) {
    setStatus("O campo Nome está vazio.");
    return;
  }

  salvarBtn.disabled = true;
  setStatus("⏳ Salvando registro...");

  try {
    const response = await fetch("/api/salvar", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await response.json();
    if (!response.ok) {
      setStatus(`✗ ${data.error || "Erro ao salvar."}`);
      return;
    }
    setStatus(data.message || "✓ Registro salvo.");
    if (data.contador !== undefined) {
      contadorEl.textContent = formatContador(data.contador);
    }
    limparCampos();
  } catch (error) {
    console.error(error);
    setStatus("✗ Erro ao salvar.");
  } finally {
    salvarBtn.disabled = false;
  }
}

async function importarPlanilha() {
  const arquivo = arquivoInput.files[0];
  if (!arquivo) {
    setStatus("Selecione um arquivo .xlsx antes de importar.");
    return;
  }

  importarBtn.disabled = true;
  setStatus("⏳ Importando planilha — aguarde...");

  const form = new FormData();
  form.append("arquivo", arquivo);
  form.append("api_key", apiKeyInput.value.trim());

  try {
    const response = await fetch("/api/importar", {
      method: "POST",
      body: form,
    });
    const data = await response.json();
    if (!response.ok) {
      setStatus(`✗ ${data.error || "Erro ao importar planilha."}`);
      return;
    }
    setStatus(data.message || "✓ Importação concluída.");
    if (data.contador !== undefined) {
      contadorEl.textContent = formatContador(data.contador);
    }
  } catch (error) {
    console.error(error);
    setStatus("✗ Erro ao importar planilha.");
  } finally {
    importarBtn.disabled = false;
    arquivoInput.value = "";
  }
}

toggleKeyBtn.addEventListener("click", () => {
  apiKeyInput.type = apiKeyInput.type === "password" ? "text" : "password";
});

processarBtn.addEventListener("click", processar);
salvarBtn.addEventListener("click", salvar);
importarBtn.addEventListener("click", importarPlanilha);
campos.status.addEventListener("change", atualizarAba);
descricaoInput.addEventListener("keydown", (event) => {
  if (event.ctrlKey && event.key === "Enter") {
    event.preventDefault();
    processar();
  }
});

preencherCampos({});
atualizarContador();
