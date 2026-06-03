/**
 * Dados de Laudos & ARTs — extraídos da planilha Excel
 * Fonte: Comissionamento_e_testes_lupa__8_.xlsx · Aba: LAUDOS
 * Última atualização planilha: 16/05/2026
 *
 * Para integrar: chame renderLaudosArts() após o DOM carregar,
 * ou substitua a função que popula a tabela "Laudos & ARTs" no seu código existente.
 */

const LAUDOS_SUMMARY = {
  totalItens: 52,
  laudosRecebidos: 0,
  aguardandoLaudo: 1,
  laudoSolicitado: 26,
  artsRecebidas: 0,
  aguardandoArt: 1,
  artSolicitada: 3,
  pctLaudosOk: 0,
  pctArtsOk: 0,
};

const LAUDOS_DATA = [
  // ── 1. ALIMENTAÇÃO E DISTRIBUIÇÃO HIDRÁULICA ──────────────────────────────
  { grupo: "1. ALIMENTAÇÃO E DISTRIBUIÇÃO HIDRÁULICA", isHeader: true },
  {
    subsistema: "1.1 Reservatório Inferior Cisterna",
    statusSistema: "REPROVADO",
    laudoNecessario: "Laudo de Limpeza e Funcionamento",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "1.2 Reservatório D'Água - Castelo",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "Laudo de Startup e Testes Funcionais",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "1.3 Bombas de Recalque",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "Laudo de Suportação de Tubulação",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "1.4 Reservatório Superior/Barrilete",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "Laudo de Startup e Testes Funcionais",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "1.5 Bombas de Pressurização",
    statusSistema: "AJUSTADO",
    laudoNecessario: "Laudo de Startup e Testes Funcionais",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "1.6 Prumadas/Shafts Hidráulicos",
    statusSistema: "REPROVADO",
    laudoNecessario: "",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "📋 SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },

  // ── 2. COLETA E DESCARTE DE ESGOTO ────────────────────────────────────────
  { grupo: "2. COLETA E DESCARTE DE ESGOTO", isHeader: true },
  {
    subsistema: "2.1 Tubulação Aparente",
    statusSistema: "REPROVADO",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "📋 SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "2.2 Bombas de Recalque - Esgoto",
    statusSistema: "AJUSTADO",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "2.3 PV's",
    statusSistema: "AJUSTADO",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "2.4 Caixa de Gordura",
    statusSistema: "AJUSTADO",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "2.5 Tubulação Enterrada",
    statusSistema: "SUB-SISTEMA APROVADO",
    laudoNecessario: "",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },

  // ── 3. COLETA, REÚSO E DESCARTE DE ÁGUAS PLUVIAIS ────────────────────────
  { grupo: "3. COLETA, REÚSO E DESCARTE DE ÁGUAS PLUVIAIS", isHeader: true },
  {
    subsistema: "3.1 Prumadas/Shafts Hidráulicos - AP",
    statusSistema: "AJUSTADO",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "3.2 Tubulações Enterradas e PV's",
    statusSistema: "AJUSTADO",
    laudoNecessario: "Laudo de Startup e Testes Funcionais",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "3.3 Bombas de Recalque Águas Pluviais",
    statusSistema: "AGENDADO",
    laudoNecessario: "",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "3.4 Bombas de Recalque Drenagem",
    statusSistema: "AGENDADO",
    laudoNecessario: "",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "3.5 Caixa de Retenção/Cacimba",
    statusSistema: "SUB-SISTEMA APROVADO",
    laudoNecessario: "",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "3.6 Tratamento de Águas Pluviais",
    statusSistema: "REPROVADO",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "3.7 Telhados, Coberturas e Rufos",
    statusSistema: "REPROVADO",
    laudoNecessario: "",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },

  // ── 4. PISCINA, SAUNA E TROCADOR DE CALOR ────────────────────────────────
  { grupo: "4. PISCINA, SAUNA E TROCADOR DE CALOR", isHeader: true },
  {
    subsistema: "4.1 Casa de Bombas da Piscina",
    statusSistema: "REPROVADO",
    laudoNecessario: "Laudo de Funcionamento da Casa de Bombas",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "❌ NÃO SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "4.2 Espelhos D'Água",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "Laudo de Startup da Sauna",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "4.3 Sauna",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "Laudo de Startup — Máquina Sauna",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "4.4 Sala de Equipamentos Sauna",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "Laudo de Funcionamento",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "4.5 Trocador de Calor",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "❌ NÃO SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "4.6 Sistema de Aquecimento da Jacuzzi",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "❌ NÃO SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "4.7 Cadeira Hidráulica",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "❌ NÃO SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },

  // ── 5. SISTEMA DE DETECÇÃO E COMBATE A INCÊNDIO ───────────────────────────
  { grupo: "5. SISTEMA DE DETECÇÃO E COMBATE A INCÊNDIO", isHeader: true },
  {
    subsistema: "5.1 Bombas de Incêndio",
    statusSistema: "REPROVADO",
    laudoNecessario: "Laudo de Startup e Testes Funcionais",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "5.2 Central de Alarme / Detectores",
    statusSistema: "REPROVADO",
    laudoNecessario: "Laudo de Testes Funcionais — Hidrantes",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "5.3 Hidrantes/Mangotinhos",
    statusSistema: "REPROVADO",
    laudoNecessario: "",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "❌ NÃO SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "5.4 Escadas Pressurizadas",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "5.5 Porta Corta Fogo",
    statusSistema: "REPROVADO",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "5.6 Extintores",
    statusSistema: "AGENDADO",
    laudoNecessario: "",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "❌ NÃO SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },

  // ── 6. ALIMENTAÇÃO E DISTRIBUIÇÃO ELÉTRICA ────────────────────────────────
  { grupo: "6. ALIMENTAÇÃO E DISTRIBUIÇÃO ELÉTRICA", isHeader: true },
  {
    subsistema: "6.1 Grupo Moto Gerador",
    statusSistema: "REPROVADO",
    laudoNecessario: "Laudo de Startup — Gerador",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "❌ NÃO SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.2 Cabine Primária",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.3 Centro de Medição",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "❌ NÃO SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.4 QGBT's",
    statusSistema: "REPROVADO",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.5 Quadros de Iluminação e Tomadas",
    statusSistema: "REPROVADO",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.6 Iluminação Externa",
    statusSistema: "AGENDADO",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.7 Shafts e Prumadas",
    statusSistema: "AGENDADO",
    laudoNecessario: "Laudo / Medição de Aterramento",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.8 Aterramento",
    statusSistema: "REPROVADO",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.9 Quadro Sinóptico",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "—",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.10 Iluminação de Emergência",
    statusSistema: "REPROVADO",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.11 Estação Recarga Veículos Elétricos",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "6.12 Subestação",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "❌ NÃO SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },

  // ── 7. CFTV, SEGURANÇA E TELEFONIA ───────────────────────────────────────
  { grupo: "7. CFTV, SEGURANÇA E TELEFONIA", isHeader: true },
  {
    subsistema: "7.1 Telefonia e Interfonia",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "7.2 CFTV",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "7.3 Segurança Perimetral",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "7.4 Portão Eletrônico",
    statusSistema: "AGENDADO",
    laudoNecessario: "",
    statusLaudo: "⏳ AGUARDANDO",
    artNecessaria: "",
    statusArt: "⏳ AGUARDANDO",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },

  // ── 8. EQUIPAMENTOS ÁREAS COMUNS E ACESSIBILIDADE ────────────────────────
  { grupo: "8. EQUIPAMENTOS ÁREAS COMUNS E ACESSIBILIDADE", isHeader: true },
  {
    subsistema: "8.1 Elevadores",
    statusSistema: "AGENDADO",
    laudoNecessario: "",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "📋 SOLICITADA",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "8.2 Plataforma Elevatória",
    statusSistema: "NÃO APLICÁVEL",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "8.3 Sistema Exaustão Churrasqueira",
    statusSistema: "AJUSTADO",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "8.4 Equipamentos de Áreas Comuns",
    statusSistema: "AGENDADO",
    laudoNecessario: "—",
    statusLaudo: "— N/A",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
  {
    subsistema: "8.5 Quadra Poliesportiva",
    statusSistema: "REPROVADO",
    laudoNecessario: "",
    statusLaudo: "📋 SOLICITADO",
    artNecessaria: "",
    statusArt: "— N/A",
    dataLaudo: "",
    dataArt: "",
    observacao: "",
  },
];

// ── Totalizadores ─────────────────────────────────────────────────────────────
const LAUDOS_TOTAIS = [
  { status: "✅  RECEBIDO / RECEBIDA", qtdLaudos: 0, pctLaudos: "0%",  qtdArts: 0, pctArts: "0%" },
  { status: "⏳  AGUARDANDO",          qtdLaudos: 1, pctLaudos: "2%",  qtdArts: 1, pctArts: "2%" },
  { status: "📋  SOLICITADO",          qtdLaudos: 26, pctLaudos: "50%", qtdArts: 3, pctArts: "6%" },
  { status: "❌  NÃO SOLICITADO",      qtdLaudos: 0, pctLaudos: "0%",  qtdArts: 9, pctArts: "17%" },
  { status: "— N/A (não requer)",      qtdLaudos: 25, pctLaudos: "48%", qtdArts: 39, pctArts: "75%" },
];

// ── Renderização ──────────────────────────────────────────────────────────────
function renderLaudosArts() {
  // 1. Preenche KPIs do cabeçalho (se existirem no DOM)
  const kpiMap = {
    "[data-laudo-total]":      LAUDOS_SUMMARY.totalItens,
    "[data-laudo-recebidos]":  LAUDOS_SUMMARY.laudosRecebidos,
    "[data-laudo-aguardando]": LAUDOS_SUMMARY.aguardandoLaudo,
    "[data-laudo-solicitado]": LAUDOS_SUMMARY.laudoSolicitado,
    "[data-art-recebidas]":    LAUDOS_SUMMARY.artsRecebidas,
    "[data-art-aguardando]":   LAUDOS_SUMMARY.aguardandoArt,
    "[data-art-solicitada]":   LAUDOS_SUMMARY.artSolicitada,
    "[data-pct-laudos]":       (LAUDOS_SUMMARY.pctLaudosOk * 100).toFixed(0) + "%",
    "[data-pct-arts]":         (LAUDOS_SUMMARY.pctArtsOk * 100).toFixed(0) + "%",
  };
  Object.entries(kpiMap).forEach(([sel, val]) => {
    document.querySelectorAll(sel).forEach((el) => (el.textContent = val));
  });

  // 2. Popula a tabela principal
  const tbody = document.querySelector("#laudos-table tbody, .laudos-table tbody, [data-laudos-tbody]");
  if (!tbody) {
    console.warn("[laudos-arts-data] Tbody alvo não encontrado. Verifique o seletor.");
    return;
  }
  tbody.innerHTML = "";

  LAUDOS_DATA.forEach((row) => {
    if (row.isHeader) {
      const tr = document.createElement("tr");
      tr.className = "laudos-group-header";
      tr.innerHTML = `<td colspan="9" class="group-label">▸  ${row.grupo}</td>`;
      tbody.appendChild(tr);
      return;
    }

    const statusClass = {
      "REPROVADO":            "status-reprovado",
      "AJUSTADO":             "status-ajustado",
      "AGENDADO":             "status-agendado",
      "NÃO APLICÁVEL":        "status-na",
      "SUB-SISTEMA APROVADO": "status-aprovado",
    }[row.statusSistema] || "";

    const laudoClass = row.statusLaudo.includes("SOLICITADO") ? "laudo-solicitado"
      : row.statusLaudo.includes("AGUARDANDO")               ? "laudo-aguardando"
      : row.statusLaudo.includes("RECEBIDO")                 ? "laudo-recebido"
      : "laudo-na";

    const artClass = row.statusArt.includes("SOLICITADA")    ? "art-solicitada"
      : row.statusArt.includes("AGUARDANDO")                 ? "art-aguardando"
      : row.statusArt.includes("RECEBIDA")                   ? "art-recebida"
      : row.statusArt.includes("NÃO SOLICITADA")             ? "art-nao-solicitada"
      : "art-na";

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="subsistema-nome">${row.subsistema}</td>
      <td><span class="badge ${statusClass}">${row.statusSistema}</span></td>
      <td>${row.laudoNecessario || "—"}</td>
      <td><span class="badge ${laudoClass}">${row.statusLaudo}</span></td>
      <td>${row.artNecessaria || "—"}</td>
      <td><span class="badge ${artClass}">${row.statusArt}</span></td>
      <td>${row.dataLaudo || "—"}</td>
      <td>${row.dataArt || "—"}</td>
      <td>${row.observacao || ""}</td>
    `;
    tbody.appendChild(tr);
  });

  console.info("[laudos-arts-data] Tabela de Laudos & ARTs populada com sucesso.");
}

// Auto-executa se o DOM já estiver pronto
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", renderLaudosArts);
} else {
  renderLaudosArts();
}
