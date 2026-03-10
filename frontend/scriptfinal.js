function toggleAgent(id) {
  document.getElementById(id).classList.toggle('show');
}

function formatCurrency(v) {
  try {
    return (v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  } catch {
    return 'R$ 0,00';
  }
}

function badgeBanco(b) {
  if (!b) return "";
  b = b.toUpperCase();
  if (b.includes("BB")) return `<span class="bank-badge badge-bb">BB</span>`;
  if (b.includes("C6")) return `<span class="bank-badge badge-c6">C6</span>`;
  return "";
}

function extractValor(motivoRaw) {
  const m = motivoRaw.match(/R\$[\s]*([\d\.,]+)/);
  if (!m) return null;

  let txt = m[1].trim();

  // Caso nº1: formato BR com milhar e decimal → 4.705,00
  if (/^\d{1,3}(\.\d{3})+,\d{2}$/.test(txt)) {
    return parseFloat(txt.replace(/\./g, "").replace(",", "."));
  }

  // Caso nº2: formato BR simples → 47,05
  if (/^\d+,\d{2}$/.test(txt)) {
    return parseFloat(txt.replace(",", "."));
  }

  // Caso nº3: formato US → 47.05
  if (/^\d+\.\d{2}$/.test(txt)) {
    return parseFloat(txt);
  }

  // Caso nº4: número inteiro → 4705
  if (/^\d+$/.test(txt)) {
    return parseFloat(txt);
  }

  // fallback final
  return parseFloat(txt.replace(",", "."));
}

/* ============================= */
/* ✅ NOVO: PARSER VALOR MANUAL   */
/* ============================= */
function parseBRLInput(v) {
  if (!v) return null;
  const s = String(v)
    .trim()
    .replace("R$", "")
    .replace(/\s+/g, "")
    .replace(/\./g, "")
    .replace(",", ".");
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
}

function getManualBaseFromEntry(btn) {
  const entry = btn.closest(".entry");
  if (!entry) return { baseNome: "", baseValorNum: null, baseValorRaw: "" };

  const baseNome = (entry.querySelector(".base-nome")?.value || "").trim();
  const baseValorRaw = (entry.querySelector(".base-valor")?.value || "").trim();
  const baseValorNum = parseBRLInput(baseValorRaw);

  return { baseNome, baseValorNum, baseValorRaw };
}

let bancoDetectado = '';
let dataConferenciaAtual = '';

document.getElementById('btnConferir').addEventListener('click', async () => {
  const pdf = document.getElementById('pdfFile').files[0];
  const excels = document.getElementById('excelFile').files;

  if (!pdf || excels.length === 0)
    return alert('Envie o PDF e pelo menos uma planilha Excel!');

  const resEl = document.getElementById('resultado');
  resEl.innerHTML = '';
  document.getElementById('progressArea').style.display = 'block';

  const fd = new FormData();

  const pdfFiles = document.getElementById('pdfFile').files;
  for (let i = 0; i < pdfFiles.length; i++) {
    fd.append('pdfs', pdfFiles[i]);
  }

  for (let i = 0; i < excels.length; i++) {
    fd.append('excels', excels[i]);
  }

  // ❌ NÃO envia mais a data para filtrar no backend

  try {
    const resp = await fetch(`${window.location.origin}/conferir_caixa`, {
      method: 'POST',
      body: fd
    });

    const dados = await resp.json();
    document.getElementById('progressArea').style.display = 'none';

    if (dados.erro) {
      resEl.innerHTML = `<div class='alert alert-danger'>${dados.erro}</div>`;
      return;
    }

    bancoDetectado = dados.banco?.toUpperCase() || 'DESCONHECIDO';

    const conf = dados.conferidos || [],
      fPdf = dados.faltando_no_pdf || [],
      fExcel = dados.faltando_no_excel || [];

    document.getElementById('totalConferidos').textContent = conf.length;
    document.getElementById('totalFaltaPdf').textContent = fPdf.length;
    document.getElementById('totalFaltaExcel').textContent = fExcel.length;

    const agentes = {};
    const add = (arr, tipo) =>
      arr.forEach((x) => {
        const ag = x.agente || 'Sem Agente';
        if (!agentes[ag])
          agentes[ag] = { conferidos: [], faltando_pdf: [], faltando_excel: [] };
        agentes[ag][tipo].push(x);
      });

    add(conf, 'conferidos');
    add(fPdf, 'faltando_pdf');
    add(fExcel, 'faltando_excel');

    let html = '';
    Object.entries(agentes).forEach(([agente, d]) => {
      const id = agente.replace(/\s+/g, '_');
      const total = d.conferidos.length + d.faltando_pdf.length + d.faltando_excel.length;
      const perc = total ? Math.round((d.conferidos.length / total) * 100) : 0;

      const circle = `
        <div class='circle-wrap'>
          <svg class='progress-ring' width='38' height='38'>
            <circle stroke='#e5e7eb' stroke-width='4' fill='transparent' r='16' cx='19' cy='19'></circle>
            <circle stroke='${perc == 100 ? '#16a34a' : '#0a66c2'}' stroke-width='4' fill='transparent' r='16' cx='19' cy='19'
              stroke-dasharray='${2 * Math.PI * 16}' stroke-dashoffset='${(1 - perc / 100) * 2 * Math.PI * 16}'></circle>
          </svg>
          <div class='circle-inner'>${perc}%</div>
        </div>`;

      function classeSetor(agente) {
        const t = (agente || "").toUpperCase();

        if (t.includes("SUPORTE ONLINE")) return "setor-suporte";
        if (t.includes("VALE VIAGENS")) return "setor-vale";
        if (t.includes("CANOA")) return "setor-canoa";
        if (t.includes("TOP VIAGENS") || t.includes("TOP")) return "setor-top";

        return "";
      }

      html += `
        <div class='agent-card ${classeSetor(agente)}'>
          <div class='agent-header' onclick='toggleAgent("${id}")'>
            <div>
              <span class='agent-name'>
                <i class='bi bi-person-circle'></i>
                ${(() => {
                  const match = agente.match(/^(.*?)(?:\s*-\s*|\s+)(SUPORTE\s+ONLINE|VALE\s+VIAGENS|TOP\s+VIAGENS|AG[ÊE]NCIA|VALE\s+AG[ÊE]NCIA)$/i);
                  if (match) {
                    const nomeBase = match[1].trim();
                    const sufixo = match[2].trim();
                    return `${nomeBase} <span class='agent-suffix'>- ${sufixo}</span>`;
                  }
                  return agente;
                })()}
              </span><br>

              <span class='agent-meta'>Conferidos: ${d.conferidos.length} • Falta PDF: ${d.faltando_pdf.length} • Falta Excel: ${d.faltando_excel.length}</span>
            </div>
            ${circle}
          </div>
          <div class='agent-content' id='${id}'>
            <div class='mt-3'>

              <!-- ================= CONFIRMADOS ================= -->
              <div class='fw-bold text-success mb-2 conferidos-titulo'>
                ✅ Conferidos (${d.conferidos.length}) —
                Total: <span class='total-conferidos'>
                  ${formatCurrency(
                    d.conferidos.reduce((acc, x) => acc + (x.valor_excel || x.valor_pdf || 0), 0)
                  )}
                </span>
              </div>

              ${d.conferidos
                .map(
                  (x, idx) => `
              <div class='entry ok' id='conferido_${id}_${idx}'>
                <div class="d-flex justify-content-between align-items-start">
                  <div>
                    <div class="fw-bold text-success mb-1">
                      ${badgeBanco(x.banco_pdf || x.banco)}
                      ${x.nome_excel || x.nome}
                    </div>

                    <div class="mt-1 ps-1">
                      <div>
                        <i class="bi bi-file-earmark-excel text-success me-1"></i>
                        <small><strong>Excel:</strong> ${x.nome_excel} — ${formatCurrency(x.valor_excel)} • ${x.hora_excel}</small>
                      </div>

                      <div>
                        <i class="bi bi-file-earmark-pdf text-danger me-1"></i>
                        <small><strong>PDF:</strong> ${badgeBanco(x.banco_pdf)} ${x.nome_pdf} — ${formatCurrency(x.valor_pdf)} • ${x.hora_pdf}</small>
                      </div>
                    </div>
                  </div>

                  <button class="btn btn-sm btn-outline-danger desmarcar-conferido"
                    data-agente="${agente}"
                    data-nome="${x.nome_excel || x.nome}"
                    data-valor="${x.valor_excel || x.valor_pdf}"
                    data-hora="${x.hora_excel || x.hora_pdf}"
                    data-banco="${x.banco_pdf || x.banco || ''}"
                    data-origem="faltando_pdf">

                  <i class="bi bi-x-circle"></i>
                </button>

                </div>
              </div>`
                )
                .join("")}

              <!-- ================= FALTANDO NO PDF ================= -->
              <div class='fw-bold text-warning mt-3 mb-2'>
                ⚠️ Faltando no PDF (${d.faltando_pdf.length})
              </div>

              ${d.faltando_pdf
                .map(
                  (x, idx) => `
              <div class='entry warn' id='faltando_${id}_${idx}'>
                <div class="d-flex justify-content-between align-items-start">
                  <div style="flex:1; padding-right:10px;">
                    <div class="fw-bold text-warning mb-1">
                      ${badgeBanco(x.banco || x.banco_pdf || x.banco_possivel)}
                      ${x.nome}
                    </div>

                    <div class="mt-1 ps-1">
                      <div>
                        <i class="bi bi-file-earmark-excel text-success me-1"></i>
                        <small><strong>Excel:</strong> ${x.nome} — ${formatCurrency(x.valor_excel ?? x.valor)} • ${x.hora}</small>
                      </div>

                      <div>
                        <i class="bi bi-file-earmark-pdf text-danger me-1"></i>
                        <small><strong>PDF:</strong> <em>não encontrado</em></small>
                      </div>
                    </div>

                    <!-- ✅ NOVO: VOCÊ DIGITA -->
                    <div class="manual-base mt-2">
                      <div class="manual-base-title">🔎 Base manual (você preenche)</div>

                      <div class="manual-base-row">
                        <input
                          type="text"
                          class="form-control form-control-sm base-nome"
                          placeholder="Nome encontrado no PDF"
                          value="${(x.base_nome || '').toString().replace(/"/g,'&quot;')}"
                        />
                        <input
                          type="text"
                          class="form-control form-control-sm base-valor"
                          placeholder="Valor (ex: 16,50)"
                          value="${(x.base_valor || '').toString().replace(/"/g,'&quot;')}"
                        />
                      </div>

                      <div class="manual-base-hint">
                        Preencha e clique ✅ para confirmar manualmente.
                      </div>
                    </div>

                  </div>

                  <button class="btn btn-sm btn-outline-success marcar-conferido"
                    data-agente="${agente}"
                    data-nome="${x.nome}"
                    data-valor="${x.valor_excel ?? x.valor}"
                    data-hora="${x.hora}"
                    data-banco="${x.banco || x.banco_pdf || x.banco_possivel || ''}">
                    <i class="bi bi-check-circle"></i>
                  </button>
                </div>
              </div>`
                )
                .join("")}

              <!-- ================= FALTANDO NO EXCEL ================= -->
              <div class='fw-bold text-danger mt-3 mb-2'>
                ❌ Faltando no Excel (${d.faltando_excel.length})
              </div>

              ${d.faltando_excel
                .map((x, idx) => `
                  <div class='entry err' id='faltando_excel_${id}_${idx}'>
                    <div class="d-flex justify-content-between align-items-start">
                      <div style="flex:1; padding-right:10px;">
                        <strong>
                          ${badgeBanco(x.banco || x.banco_pdf || x.banco_excel)}
                          ${x.nome || x.nome_pdf || "(sem nome)"}
                        </strong>

                        <div class="mt-1">
                          <div>
                            <i class="bi bi-file-earmark-excel text-success"></i>
                            <small><strong>Excel:</strong> <em>não encontrado</em></small>
                          </div>

                          <div>
                            <i class="bi bi-file-earmark-pdf text-danger"></i>
                            <small>
                              <strong>PDF:</strong> ${formatCurrency(x.valor)} • ${x.hora}
                            </small>
                          </div>
                        </div>
                      </div>

                      <button
                        type="button"
                        class="btn btn-sm btn-outline-danger excluir-faltando-excel"
                        title="Excluir este item"
                        data-nome="${(x.nome || x.nome_pdf || '').replace(/"/g, '&quot;')}"
                      >
                        <i class="bi bi-trash"></i>
                      </button>
                    </div>
                  </div>`)
                .join("")}

            </div>
          </div>
        </div>`;
    });

    resEl.innerHTML = html;
    recalcAll();

    // 🔧 Ajuste visual para o agente "Sem Agente"
    document.querySelectorAll('.agent-card').forEach(card => {
      const agentName = card.querySelector('.agent-name')?.textContent.trim().toLowerCase() || '';

      if (agentName !== 'sem agente') {
        card.querySelectorAll('.fw-bold.text-danger, .entry.err').forEach(el => el.remove());
      } else {
        card.querySelectorAll('.fw-bold.text-success, .entry.ok, .fw-bold.text-warning, .entry.warn').forEach(el => el.remove());
        card.querySelectorAll('.fw-bold.text-danger').forEach(el => el.remove());
        card.querySelectorAll('.entry.err .titulo').forEach(el => el.remove());

        const faltandoExcelCount = card.querySelectorAll('.entry.err').length;

        const header = card.querySelector('.agent-header');
        if (header) {
          header.innerHTML = `
            <div class="fw-bold" style="color:#a31515; font-size:1.1rem;">
              ❌ FALTANDO EXCEL : ${faltandoExcelCount}
            </div>
          `;
        }

        const meta = card.querySelector('.agent-meta');
        if (meta) meta.remove();
      }
    });

    // 🟢 Marcar item como conferido
    document.querySelectorAll('.marcar-conferido').forEach(btn => {
      btn.addEventListener('click', () => {
        moverItem(btn, 'faltando', 'conferido');
      });
    });

    // 🔴 Desmarcar item (voltar para faltando no PDF)
    document.querySelectorAll('.desmarcar-conferido').forEach(btn => {
      btn.addEventListener('click', () => {
        moverItem(btn, 'conferido', 'faltando');
      });
    });

    // 🗑️ Excluir item de "Faltando no Excel"
    document.querySelectorAll('.excluir-faltando-excel').forEach(btn => {
      btn.addEventListener('click', () => {
        excluirFaltandoExcel(btn);
      });
    });

    function updateAllCountsAndRender(agentIdToFocus = null) {
      try {
        const agentCards = Array.from(document.querySelectorAll('.agent-card'));
        let totalConferidosGlob = 0;
        let totalFaltaPdfGlob = 0;
        let totalFaltaExcelGlob = 0;
        let totalValorConferidosGlob = 0;

        agentCards.forEach(card => {
          const agentContent = card.querySelector('.agent-content');
          const agentId = agentContent ? agentContent.id : null;

          const conferidosEls = card.querySelectorAll('.entry.ok');
          const faltandoEls = card.querySelectorAll('.entry.warn');
          const faltaExcelEls = card.querySelectorAll('.entry.err');

          const conferidosCount = conferidosEls.length;
          const faltandoCount = faltandoEls.length;
          const faltaExcelCount = faltaExcelEls.length;

          let totalValor = 0;
          conferidosEls.forEach(el => {
            const txt = el.innerText || '';
            const m = txt.match(/R\$[\s]*([\d\.\,]+)/);
            if (m && m[1]) {
              const numStr = m[1].trim().replace(/\./g, '').replace(',', '.');
              totalValor += parseFloat(numStr) || 0;
            } else {
              const dv = el.dataset && el.dataset.valor;
              if (dv) totalValor += parseFloat(dv) || 0;
            }
          });

          const metaEl = card.querySelector('.agent-meta');
          if (metaEl) {
            metaEl.textContent = `Conferidos: ${conferidosCount} • Falta PDF: ${faltandoCount} • Falta Excel: ${faltaExcelCount}`;
          }

          const confTituloEl = card.querySelector('.conferidos-titulo');
          if (confTituloEl) {
            confTituloEl.innerHTML = `✅ Conferidos (${conferidosCount}) — Total: <span class='total-conferidos'>${formatCurrency(totalValor)}</span>`;
          }

          const faltTituloEl = card.querySelector('.fw-bold.text-warning');
          if (faltTituloEl) {
            faltTituloEl.innerHTML = `⚠️ Faltando no PDF (${faltandoCount})`;
          }

          const isSemAgenteHeader = (card.querySelector('.agent-header')?.innerText || '').toUpperCase().includes('FALTANDO EXCEL');
          if (isSemAgenteHeader) {
            const headerDiv = card.querySelector('.agent-header div');
            if (headerDiv) {
              headerDiv.innerHTML = `❌ FALTANDO EXCEL : ${faltaExcelCount}`;
            }
          }

          const totalItens = Math.max(1, conferidosCount + faltandoCount + faltaExcelCount);
          const perc = Math.round((conferidosCount / totalItens) * 100);
          const circles = card.querySelectorAll('.progress-ring circle');
          const circle = circles.length > 1 ? circles[1] : circles[0];
          const inner = card.querySelector('.circle-inner');
          if (circle && inner) {
            const r = parseFloat(circle.getAttribute('r')) || 16;
            const circ = 2 * Math.PI * r;
            const offset = ((1 - perc / 100) * circ).toFixed(2);
            circle.style.transition = 'stroke-dashoffset 0.3s ease, stroke 0.3s ease';
            circle.setAttribute('stroke-dashoffset', offset);
            circle.setAttribute('stroke', perc === 100 ? '#16a34a' : '#0a66c2');
            inner.textContent = `${perc}%`;
          }

          totalConferidosGlob += conferidosCount;
          totalFaltaPdfGlob += faltandoCount;
          totalFaltaExcelGlob += faltaExcelCount;
          totalValorConferidosGlob += totalValor;
        });

        const totalConferidosEl = document.getElementById('totalConferidos');
        const totalFaltaPdfEl = document.getElementById('totalFaltaPdf');
        const totalFaltaExcelEl = document.getElementById('totalFaltaExcel');

        if (totalConferidosEl) totalConferidosEl.textContent = totalConferidosGlob;
        if (totalFaltaPdfEl) totalFaltaPdfEl.textContent = totalFaltaPdfGlob;
        if (totalFaltaExcelEl) totalFaltaExcelEl.textContent = totalFaltaExcelGlob;

        const topoTotalSpan = document.querySelector('.total-conferidos-topo');
        if (topoTotalSpan) topoTotalSpan.textContent = formatCurrency(totalValorConferidosGlob);

        if (agentIdToFocus) {
          const agentContent = document.getElementById(agentIdToFocus);
          if (agentContent) {
            const card = agentContent.closest('.agent-card');
            if (card) {
              card.style.transition = 'box-shadow 0.2s';
              card.style.boxShadow = '0 0 0 3px rgba(10,102,194,0.08)';
              setTimeout(() => card.style.boxShadow = '', 350);
            }
          }
        }

        return true;
      } catch (err) {
        console.warn('updateAllCountsAndRender erro:', err);
        return false;
      }
    }

    function recalcAndRenderAgent(agenteId) {
      updateAllCountsAndRender(agenteId);
    }

    function normalizeText(s) {
      return String(s || "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim();
    }

    function parseCurrencyFromText(txt) {
      if (!txt) return null;
      const m = String(txt).match(/R\$[\s]*([\d\.\,]+)/);
      if (!m) return null;
      const numStr = m[1].replace(/\./g, "").replace(",", ".");
      const n = parseFloat(numStr);
      return Number.isFinite(n) ? n : null;
    }

    function removeMatchingFromFaltandoExcel({ nomeToMatch, valorToMatch }) {
      const nomeNorm = normalizeText(nomeToMatch);
      const valorNum = typeof valorToMatch === "number"
        ? valorToMatch
        : parseFloat(valorToMatch) || null;

      const errItems = Array.from(document.querySelectorAll('.entry.err'));
      let removidos = 0;

      errItems.forEach(el => {
        try {
          const strong = el.querySelector('strong');
          const nomeErr = normalizeText(strong ? strong.textContent : "");
          const valorErr = parseCurrencyFromText(el.innerText);

          const matchNome =
            nomeNorm &&
            (nomeErr.includes(nomeNorm) || nomeNorm.includes(nomeErr));

          const matchValor =
            valorNum != null && valorErr != null
              ? Math.abs(valorErr - valorNum) < 0.05
              : true;

          if (matchNome && matchValor) {
            const agentContent = el.closest(".agent-content");
            const agentId = agentContent ? agentContent.id : null;

            el.style.transition = "opacity 0.35s ease, transform 0.35s ease";
            el.style.opacity = "0";
            el.style.transform = "translateX(-16px)";

            setTimeout(() => {
              el.remove();
              if (agentId) recalcAndRenderAgent(agentId);
            }, 360);

            removidos++;
          }
        } catch (err) {
          console.warn("Erro ao remover faltando_excel:", err);
        }
      });

      if (removidos > 0) {
        setTimeout(() => {
          document.getElementById("totalFaltaExcel").textContent =
            document.querySelectorAll(".entry.err").length;
        }, 400);
      }

      return removidos;
    }

    function recalcAll() {
      const totalConferidos = document.querySelectorAll(".entry.ok").length;
      const totalFaltaPDF = document.querySelectorAll(".entry.warn").length;
      const totalFaltaExcel = document.querySelectorAll(".entry.err").length;

      document.getElementById("totalConferidos").textContent = totalConferidos;
      document.getElementById("totalFaltaPdf").textContent = totalFaltaPDF;
      document.getElementById("totalFaltaExcel").textContent = totalFaltaExcel;

      document.querySelectorAll(".agent-content").forEach(ac => {
        const agentId = ac.id;
        if (agentId) recalcAndRenderAgent(agentId);
      });
    }

    function excluirFaltandoExcel(btn) {
      const item = btn.closest('.entry.err');
      if (!item) return;

      item.style.transition = 'opacity 0.25s ease, transform 0.25s ease';
      item.style.opacity = '0';
      item.style.transform = 'translateX(12px)';

      setTimeout(() => {
        item.remove();
        recalcAll();
      }, 260);
    }

    // 🔄 mover itens (com base manual persistente)
    function moverItem(btn, origem, destino) {

      const origemReal = btn.dataset.origem || origem;
      const voltarExcel = btn.dataset.voltarexcel === "1";
      const nome = btn.dataset.nome?.trim() || "";
      let valor = parseFloat(btn.dataset.valor || 0);
      const hora = btn.dataset.hora || "(sem hora)";
      const agente = btn.dataset.agente;
      const banco = btn.dataset.banco || "";
      const motivoSalvo = decodeURIComponent(btn.dataset.motivo || "");
      const agenteId = agente.replace(/\s+/g, "_");

      // ✅ base manual salva no botão (quando desmarca/volta)
      const baseNomeSaved = decodeURIComponent(btn.dataset.base_nome || "");
      const baseValorSaved = decodeURIComponent(btn.dataset.base_valor || "");

      const card = btn.closest(`.entry.${origem === "conferido" ? "ok" : "warn"}`);
      let motivoRaw = "";

      // ✅ se for marcar (origem warn), pega o que você digitou
      const manual = getManualBaseFromEntry(btn);

      if (card) {
        motivoRaw =
          card.querySelector(".text-muted small")?.textContent ||
          motivoSalvo ||
          "";
        card.remove();
      }

      // ============================================================
      // DESTINO: CONFERIDO
      // ============================================================
      if (destino === "conferido") {

        let detalheManual = "confirmado manualmente";

        // ✅ pega base manual do input (se vazio, tenta reaproveitar do dataset antigo)
        const baseNome = manual.baseNome || baseNomeSaved || null;

        // valor base: se digitou usa, se não digitou tenta dataset salvo, se não tenta motivo antigo
        const baseValorNumFinal =
          (manual.baseValorNum != null) ? manual.baseValorNum :
          (parseBRLInput(baseValorSaved) != null) ? parseBRLInput(baseValorSaved) :
          extractValor(motivoRaw);

        if (baseNome && baseValorNumFinal != null) {
          detalheManual =
            `confirmado manualmente (baseado em ${baseNome} — R$${baseValorNumFinal
              .toFixed(2)
              .replace(".", ",")})`;
        }

        const confContainer = document.querySelector(`#${agenteId} .fw-bold.text-success`);

        if (confContainer) {
          const novo = document.createElement("div");
          novo.className = "entry ok";

          // REMOVE DO FALTANDO EXCEL (se tiver base manual completa)
          let removidosExcel = 0;
          if (baseNome && baseValorNumFinal != null) {
            removidosExcel = removeMatchingFromFaltandoExcel({
              nomeToMatch: baseNome,
              valorToMatch: baseValorNumFinal
            });
          }

          const baseValorParaSalvar = (manual.baseValorRaw || baseValorSaved || "");

          novo.innerHTML = `
          <div class="d-flex justify-content-between align-items-start">
            <div>
              <div class="fw-bold text-success mb-1">
                ${badgeBanco(banco)} ${nome}
              </div>
              <div class="mt-1 ps-1">
                <div>
                  <i class="bi bi-file-earmark-excel text-success me-1"></i>
                  <small><strong>Excel:</strong> ${nome} — ${formatCurrency(valor)} • ${hora}</small>
                </div>
                <div>
                  <i class="bi bi-file-earmark-pdf text-danger me-1"></i>
                  <small><strong>PDF:</strong> <em>${detalheManual}</em></small>
                </div>
              </div>
            </div>

            <button class="btn btn-sm btn-outline-danger desmarcar-conferido"
              data-agente="${agente}"
              data-nome="${nome}"
              data-valor="${valor}"
              data-hora="${hora}"
              data-banco="${banco}"
              data-origem="${origemReal}"
              data-motivo="${encodeURIComponent(motivoRaw)}"
              data-voltarexcel="${removidosExcel > 0 ? "1" : "0"}"
              data-base_nome="${encodeURIComponent(baseNome || "")}"
              data-base_valor="${encodeURIComponent(baseValorParaSalvar)}"
            >
              <i class="bi bi-x-circle"></i>
            </button>
          </div>`;

          novo.querySelector(".desmarcar-conferido").addEventListener("click", () =>
            moverItem(novo.querySelector(".desmarcar-conferido"), "conferido", "faltando")
          );

          confContainer.insertAdjacentElement("afterend", novo);

          recalcAll();
        }

        return;
      }

      // ============================================================
      // DESTINO: VOLTAR (DESMARCAR)
      // ============================================================

      // 1️⃣ VOLTAR PARA FALTANDO EXCEL (se foi removido lá)
      if (voltarExcel) {

        // tenta usar base manual salva (pra voltar certinho)
        const nomeFinal = baseNomeSaved || nome;
        const valorFinal = parseBRLInput(baseValorSaved) ?? valor;

        const semAgenteCard = Array.from(document.querySelectorAll(".agent-card"))
          .find(card =>
            card.querySelector(".agent-header")?.innerText
              .toUpperCase()
              .includes("FALTANDO EXCEL")
          );

        if (semAgenteCard) {

          const novo = document.createElement("div");
          novo.className = "entry err";

          novo.innerHTML = `
            <div class="d-flex justify-content-between align-items-start">
              <div style="flex:1; padding-right:10px;">
                <strong>${badgeBanco(banco)} ${nomeFinal}</strong>
                <div class="mt-1">
                  <div>
                    <i class="bi bi-file-earmark-excel text-success"></i>
                    <small><strong>Excel:</strong> <em>não encontrado</em></small>
                  </div>
                  <div>
                    <i class="bi bi-file-earmark-pdf text-danger"></i>
                    <small><strong>PDF:</strong> ${formatCurrency(valorFinal)} • ${hora}</small>
                  </div>
                </div>
              </div>

              <button
                type="button"
                class="btn btn-sm btn-outline-danger excluir-faltando-excel"
                title="Excluir este item"
                data-nome="${(nomeFinal || '').replace(/"/g, '&quot;')}"
              >
                <i class="bi bi-trash"></i>
              </button>
            </div>`;

          semAgenteCard.querySelector(".agent-content").appendChild(novo);

          const btnExcluirNovo = novo.querySelector(".excluir-faltando-excel");
          if (btnExcluirNovo) {
            btnExcluirNovo.addEventListener("click", () => {
              excluirFaltandoExcel(btnExcluirNovo);
            });
          }

          const localCount = semAgenteCard.querySelectorAll(".entry.err").length;
          const tituloSemAgente = semAgenteCard.querySelector(".fw-bold");
          if (tituloSemAgente) {
            tituloSemAgente.innerHTML = `❌ FALTANDO EXCEL : ${localCount}`;
          }

          document.getElementById("totalFaltaExcel").textContent =
            document.querySelectorAll(".entry.err").length;

          recalcAndRenderAgent(agenteId);
        }
      }

      // 2️⃣ VOLTAR PARA FALTANDO PDF SEMPRE (com inputs preenchidos)
      const faltContainer = document.querySelector(`#${agenteId} .fw-bold.text-warning`);

      if (faltContainer) {
        const novo = document.createElement("div");
        novo.className = "entry warn";

        novo.innerHTML = `
        <div class="d-flex justify-content-between align-items-start">
          <div style="flex:1; padding-right:10px;">
            <div class="fw-bold text-warning mb-1">
              ${badgeBanco(banco)} ${nome}
            </div>

            <div class="mt-1 ps-1">
              <div>
                <i class="bi bi-file-earmark-excel text-success me-1"></i>
                <small><strong>Excel:</strong> ${nome} — ${formatCurrency(valor)} • ${hora}</small>
              </div>
              <div>
                <i class="bi bi-file-earmark-pdf text-danger me-1"></i>
                <small><strong>PDF:</strong> <em>não encontrado</em></small>
              </div>
            </div>

            <div class="manual-base mt-2">
              <div class="manual-base-title">🔎 Base manual (você preenche)</div>
              <div class="manual-base-row">
                <input type="text" class="form-control form-control-sm base-nome"
                  placeholder="Nome encontrado no PDF"
                  value="${(baseNomeSaved || '').toString().replace(/"/g,'&quot;')}" />
                <input type="text" class="form-control form-control-sm base-valor"
                  placeholder="Valor (ex: 16,50)"
                  value="${(baseValorSaved || '').toString().replace(/"/g,'&quot;')}" />
              </div>
              <div class="manual-base-hint">Preencha e clique ✅ para confirmar manualmente.</div>
            </div>

            ${motivoRaw ? `
            <div class="text-muted mt-1">
              <small>${motivoRaw}</small>
            </div>` : ""}
          </div>

          <button class="btn btn-sm btn-outline-success marcar-conferido"
            data-agente="${agente}"
            data-nome="${nome}"
            data-valor="${valor}"
            data-hora="${hora}"
            data-banco="${banco}"
            data-origem="${origemReal}"
            data-motivo="${encodeURIComponent(motivoRaw)}"
            data-voltarexcel="${voltarExcel ? "1" : "0"}"
            data-base_nome="${encodeURIComponent(baseNomeSaved || '')}"
            data-base_valor="${encodeURIComponent(baseValorSaved || '')}"
          >
            <i class="bi bi-check-circle"></i>
          </button>
        </div>`;

        faltContainer.insertAdjacentElement("afterend", novo);

        novo.querySelector(".marcar-conferido").addEventListener("click", () =>
          moverItem(novo.querySelector(".marcar-conferido"), "faltando", "conferido")
        );
      }

      recalcAndRenderAgent(agenteId);
    }

  } catch (e) {
    document.getElementById('progressArea').style.display = 'none';
    resEl.innerHTML = `<div class='alert alert-danger'>Erro: ${e.message}</div>`;
  }
});

document.getElementById('btnLimpar').addEventListener('click', () => {
  document.getElementById('pdfFile').value = '';
  document.getElementById('excelFile').value = '';
  document.getElementById('dataFiltro').value = '';
  document.getElementById('resultado').innerHTML = '';
  document.getElementById('totalConferidos').textContent = '0';
  document.getElementById('totalFaltaPdf').textContent = '0';
  document.getElementById('totalFaltaExcel').textContent = '0';
  dataConferenciaAtual = '';
});
function formatarDataBR(dataStr) {
  if (!dataStr) return '';
  const s = String(dataStr).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const [ano, mes, dia] = s.split('-');
    return `${dia}/${mes}/${ano}`;
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    return s;
  }

  if (/^\d{2}-\d{2}-\d{4}$/.test(s)) {
    const [dia, mes, ano] = s.split('-');
    return `${dia}/${mes}/${ano}`;
  }

  return s;
}

function extrairDataRelatorio() {
  if (dataConferenciaAtual) {
    return formatarDataBR(dataConferenciaAtual);
  }

  const dataFiltro = document.getElementById('dataFiltro')?.value?.trim();
  if (dataFiltro) {
    return formatarDataBR(dataFiltro);
  }

  return new Date().toLocaleDateString('pt-BR');
}

function parseValorBRL(texto) {
  if (!texto) return 0;
  const m = String(texto).match(/R\$[\s]*([\d\.\,]+)/);
  if (!m) return 0;
  return parseFloat(m[1].replace(/\./g, '').replace(',', '.')) || 0;
}
function formatarDataRelatorioDoInput() {
  const valor = document.getElementById('dataFiltro')?.value?.trim();

  if (!valor) return 'Não informada';

  if (/^\d{4}-\d{2}-\d{2}$/.test(valor)) {
    const [ano, mes, dia] = valor.split('-');
    return `${dia}/${mes}/${ano}`;
  }

  return valor;
}

document.getElementById('btnExport').addEventListener('click', async () => {
  const totalC = document.getElementById('totalConferidos').textContent;
  const totalP = document.getElementById('totalFaltaPdf').textContent;
  const totalE = document.getElementById('totalFaltaExcel').textContent;

  const dataRelatorio = formatarDataRelatorioDoInput();
  const nomeBase = `ConferenciaCaixa_${bancoDetectado || 'DESCONHECIDO'}_${dataRelatorio.replace(/[\/\s]/g, '-')}`;

  const agentes = Array.from(document.querySelectorAll('.agent-card'));

  const dadosResumo = agentes.map(card => {
    const nome = card.querySelector('.agent-name')?.textContent.trim() || 'Sem Agente';

    const meta = card.querySelector('.agent-meta')?.textContent || '';
    const matchC = meta.match(/Conferidos:\s*(\d+)/);
    const matchP = meta.match(/Falta PDF:\s*(\d+)/);
    const matchE = meta.match(/Falta Excel:\s*(\d+)/);

    let conferidos = parseInt(matchC?.[1] || 0);
    let faltaPdf = parseInt(matchP?.[1] || 0);
    let faltaExcel = parseInt(matchE?.[1] || 0);

    let totalValor = 'R$ 0,00';
    let perc = card.querySelector('.circle-inner')?.textContent.trim() || '0%';

    const headerTxt = card.querySelector('.agent-header')?.innerText?.trim() || '';
    const isSemAgente =
      nome === 'Sem Agente' ||
      headerTxt.toUpperCase().includes('FALTANDO EXCEL');

    if (isSemAgente) {
      const itensErr = Array.from(card.querySelectorAll('.entry.err'));
      faltaExcel = itensErr.length;
      conferidos = 0;
      faltaPdf = 0;

      let soma = 0;
      itensErr.forEach(el => {
        soma += parseValorBRL(el.innerText);
      });

      totalValor = formatCurrency(soma);
      perc = '0%';
    } else {
      const totalSpan = card.querySelector('.total-conferidos');
      totalValor = totalSpan ? totalSpan.textContent.trim() : 'R$ 0,00';
    }

    return { nome, conferidos, faltaPdf, faltaExcel, totalValor, perc };
  });

  const linhas = dadosResumo.map(a => `
    <tr>
      <td>${a.nome}</td>
      <td style="text-align:center;">${a.conferidos}</td>
      <td style="text-align:center;">${a.faltaPdf}</td>
      <td style="text-align:center;">${a.faltaExcel}</td>
      <td style="text-align:right;">${a.totalValor}</td>
      <td style="text-align:center;">${a.perc}</td>
    </tr>`).join('');

  const wrapper = document.createElement('div');
  wrapper.innerHTML = `
    <style>
      * { font-family: Arial, sans-serif !important; color:#111 !important; }
      h2 { margin: 0; }
      table { width: 100%; border-collapse: collapse; font-size: 12px; }
      th, td { border: 1px solid #ccc; padding: 6px; }
      th { background: #0a66c2; color: #fff !important; }
      td { background: #fff; }
      .topo { text-align:center; margin-bottom:10px; }
      .linha { margin:2px 0; font-size:12px; }
      hr { margin: 10px 0; }
    </style>

    <div class="topo">
      <h2>📊 Resumo de Conferência de Caixa</h2>
      <div class="linha">Banco: <strong>${bancoDetectado}</strong> • Data: <strong>${dataRelatorio}</strong></div>
      <div class="linha">✅ Conferidos: ${totalC} • ⚠️ Falta PDF: ${totalP} • ❌ Falta Excel: ${totalE}</div>
      <hr>
    </div>

    <table>
      <thead>
        <tr>
          <th>Agente</th>
          <th>Conferidos</th>
          <th>Falta PDF</th>
          <th>Falta Excel</th>
          <th>Total (R$)</th>
          <th>%</th>
        </tr>
      </thead>
      <tbody>${linhas}</tbody>
    </table>

    <div style="text-align:center; margin-top:10px; font-size:10px; color:#444 !important;">
      © ${new Date().getFullYear()} Conferência de Caixa — Desenvolvido por <strong>Gilmario Lima</strong>
    </div>
  `;

  const optFinal = {
    margin: [0.3, 0.3, 0.4, 0.3],
    filename: nomeBase + '_Resumo.pdf',
    html2canvas: { scale: 1.2, useCORS: true, backgroundColor: '#ffffff' },
    jsPDF: { unit: 'in', format: 'a4', orientation: 'portrait' },
    pagebreak: { mode: ['css', 'legacy'] }
  };

  html2pdf().set(optFinal).from(wrapper).save();
});
