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

let bancoDetectado = '';

document.getElementById('btnConferir').addEventListener('click', async () => {
  const pdf = document.getElementById('pdfFile').files[0];
  const excels = document.getElementById('excelFile').files;
  const dataFiltro = document.getElementById('dataFiltro').value;
  if (!pdf || excels.length === 0)
    return alert('Envie o PDF e pelo menos uma planilha Excel!');

  const resEl = document.getElementById('resultado');
  resEl.innerHTML = '';
  document.getElementById('progressArea').style.display = 'block';

  const fd = new FormData();
  fd.append('pdf', pdf);
  for (let i = 0; i < excels.length; i++) fd.append('excels', excels[i]);
  fd.append('data', dataFiltro);

  try {
    const backendURL = window.location.origin; // Detecta o domínio atual (Render)
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

      html += `
        <div class='agent-card'>
          <div class='agent-header' onclick='toggleAgent("${id}")'>
            <div>
              <span class='agent-name'>
                <i class='bi bi-person-circle'></i>
                ${(() => {
                  // separa sufixos conhecidos e colore
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
              <div class='fw-bold text-success mb-2 conferidos-titulo'>
                ✅ Conferidos (${d.conferidos.length}) — 
                Total: <span class='total-conferidos'>${formatCurrency(
                  d.conferidos.reduce((acc, x) => acc + (x.valor_excel || x.valor_pdf || 0), 0)
                )}</span>
              </div>
              ${d.conferidos.map((x, idx) => `
                <div class='entry ok' id='conferido_${id}_${idx}'>
                  <div class="d-flex justify-content-between align-items-start">
                    <div>
                      <div class="fw-bold text-success mb-1">${x.nome_excel || x.nome}</div>
                      <div class="mt-1 ps-1">
                        <div><i class="bi bi-file-earmark-excel text-success me-1"></i>
                          <small><strong>Excel:</strong> ${x.nome_excel || '-'} — ${formatCurrency(x.valor_excel)} • ${x.hora_excel || '(sem hora)'}</small>
                        </div>
                        <div><i class="bi bi-file-earmark-pdf text-danger me-1"></i>
                          <small><strong>PDF:</strong> ${x.nome_pdf || '-'} — ${formatCurrency(x.valor_pdf)} • ${x.hora_pdf || '(sem hora)'}</small>
                        </div>
                      </div>
                    </div>
                    <!-- 🔴 Botão para desmarcar conferido -->
                    <button class="btn btn-sm btn-outline-danger desmarcar-conferido"
                            data-agente="${agente}"
                            data-nome="${x.nome_excel || x.nome}"
                            data-valor="${x.valor_excel || x.valor_pdf || 0}"
                            data-hora="${x.hora_excel || x.hora_pdf || ''}">
                      <i class="bi bi-x-circle"></i>
                    </button>
                  </div>
                </div>
              `).join('')}

              <div class='fw-bold text-warning mt-3 mb-2'>⚠️ Faltando no PDF (${d.faltando_pdf.length})</div>
              ${d.faltando_pdf.map((x, idx) => `
                <div class='entry warn' id='faltando_${id}_${idx}'>
                  <div class="d-flex justify-content-between align-items-start">
                    <div>
                      <div class="fw-bold text-warning mb-1">${x.nome}</div>
                      <div class="mt-1 ps-1">
                        <div><i class="bi bi-file-earmark-excel text-success me-1"></i>
                          <small><strong>Excel:</strong> ${x.nome || '-'} — ${formatCurrency(x.valor_excel ?? x.valor)} • ${x.hora || '(sem hora)'}</small>
                        </div>
                        <div><i class="bi bi-file-earmark-pdf text-danger me-1"></i>
                          <small><strong>PDF:</strong> <em>não encontrado</em></small>
                        </div>
                      </div>
                      <div class="text-muted mt-1"><small>💬 ${x.motivo || 'Sem motivo registrado.'}</small></div>
                    </div>

                    <!-- 🟢 Botão para marcar como conferido -->
                    <button class="btn btn-sm btn-outline-success marcar-conferido"
                            data-agente="${agente}"
                            data-nome="${x.nome}"
                            data-valor="${x.valor_excel ?? x.valor}"
                            data-hora="${x.hora}">
                      <i class="bi bi-check-circle"></i>
                    </button>
                  </div>
                </div>
              `).join('')}



              <div class='fw-bold text-danger mt-3 mb-2'>❌ Faltando no Excel (${d.faltando_excel.length})</div>
              ${d.faltando_excel.map(x => `
                <div class='entry err'>
                  <strong>${x.nome}</strong>
                  <div class="mt-1">
                    <div><i class="bi bi-file-earmark-excel text-success"></i>
                      <small><strong>Excel:</strong> <em>não encontrado</em></small>
                    </div>
                    <div><i class="bi bi-file-earmark-pdf text-danger"></i>
                      <small><strong>PDF:</strong> ${formatCurrency(x.valor)} • ${x.hora || '(sem hora)'}</small>
                    </div>
                  </div>
                </div>`).join('')}
            </div>
          </div>
        </div>`;
    });

    resEl.innerHTML = html;

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



    // Recalcula contadores, soma total dos "conferidos" e atualiza o círculo de progresso
    function recalcAndRenderAgent(agenteId) {
      try {
        // agenteId exemplo: "GILMARIO_LIMA" — corresponde ao id do .agent-content
        const agentContent = document.getElementById(agenteId);
        if (!agentContent) return;

        const agentCard = agentContent.closest('.agent-card');
        if (!agentCard) return;

        // contar entradas
        const conferidosEls = agentCard.querySelectorAll('.entry.ok');
        const faltandoEls = agentCard.querySelectorAll('.entry.warn');
        const faltaExcelEls = agentCard.querySelectorAll('.entry.err');

        const conferidosCount = conferidosEls.length;
        const faltandoCount = faltandoEls.length;
        const faltaExcelCount = faltaExcelEls.length;

        // somar valores dos conferidos: tenta extrair R$ XX,XX do texto de cada entry.ok
        let totalValor = 0;
        conferidosEls.forEach(el => {
          // procura primeiro "R$ x.xxx,xx" no texto do elemento
          const txt = el.innerText || '';
          const m = txt.match(/R\$[\s]*([\d\.\,]+)/);
          if (m && m[1]) {
            const numStr = m[1].trim().replace(/\./g, '').replace(',', '.'); // 1.234,56 -> 1234.56
            const n = parseFloat(numStr) || 0;
            totalValor += n;
          } else {
            // fallback: checa atributo data-valor se existir
            const dv = el.dataset && el.dataset.valor;
            if (dv) totalValor += parseFloat(dv) || 0;
          }
        });

        // atualizar meta (linha pequena abaixo do nome do agente)
        const metaEl = agentCard.querySelector('.agent-meta');
        if (metaEl) {
          metaEl.textContent = `Conferidos: ${conferidosCount} • Falta PDF: ${faltandoCount} • Falta Excel: ${faltaExcelCount}`;
        }

        // atualizar título de conferidos e total
        const confTituloEl = agentCard.querySelector('.conferidos-titulo');
        if (confTituloEl) {
          const totalFormatted = formatCurrency(totalValor);
          confTituloEl.innerHTML = `✅ Conferidos (${conferidosCount}) — Total: <span class='total-conferidos'>${totalFormatted}</span>`;
        }

        // atualizar contador amarelo (faltando)
        const faltTituloEl = agentCard.querySelector('.fw-bold.text-warning');
        if (faltTituloEl) {
          faltTituloEl.innerHTML = `⚠️ Faltando no PDF (${faltandoCount})`;
        }

        // atualizar círculo de progresso (percentual)
        const totalItens = Math.max(1, conferidosCount + faltandoCount + faltaExcelCount);
        const perc = Math.round((conferidosCount / totalItens) * 100);

        // pega o segundo circle (o visível) e o .circle-inner
        const circles = agentCard.querySelectorAll('.progress-ring circle');
        const circle = circles.length > 1 ? circles[1] : circles[0];
        const inner = agentCard.querySelector('.circle-inner');

        if (circle && inner) {
          const r = parseFloat(circle.getAttribute('r')) || 16;
          const circ = 2 * Math.PI * r;
          const offset = ((1 - perc / 100) * circ).toFixed(2);
          circle.style.transition = 'stroke-dashoffset 0.3s ease, stroke 0.3s ease';
          circle.setAttribute('stroke-dashoffset', offset);
          circle.setAttribute('stroke', perc === 100 ? '#16a34a' : '#0a66c2');
          inner.textContent = `${perc}%`;
        }
      } catch (err) {
        console.warn('recalcAndRenderAgent erro:', err);
      }
    }


    // 🔄 Função geral para mover itens entre listas
    function moverItem(btn, origem, destino) {
      const nome = btn.dataset.nome;
      const valor = parseFloat(btn.dataset.valor || 0);
      const hora = btn.dataset.hora || '(sem hora)';
      const agente = btn.dataset.agente;
      const agenteId = agente.replace(/\s+/g, '_');

      // Remove o card da origem
      const card = btn.closest(`.entry.${origem === 'conferido' ? 'ok' : 'warn'}`);
      if (card) card.remove();

      // Seleciona o container e contadores
      const meta = document.querySelector(`#${agenteId} .agent-meta`);
      const confTitulo = document.querySelector(`#${agenteId} .conferidos-titulo`);
      const totalSpan = confTitulo?.querySelector('.total-conferidos');
      const confCountMatch = confTitulo?.textContent.match(/Conferidos\s*\((\d+)\)/);
      const confCount = confCountMatch ? parseInt(confCountMatch[1]) : 0;
      const totalValor = parseFloat(totalSpan?.textContent.replace(/[^\d,.-]/g, '').replace(',', '.') || 0);

      const faltandoTitulo = document.querySelector(`#${agenteId} .fw-bold.text-warning`);
      const faltandoMatch = faltandoTitulo?.textContent.match(/\((\d+)\)/);
      const faltandoCount = faltandoMatch ? parseInt(faltandoMatch[1]) : 0;

      if (destino === 'conferido') {
        // ➕ Adiciona ao conferido
        const confContainer = document.querySelector(`#${agenteId} .fw-bold.text-success`);
        if (confContainer) {
          const novo = document.createElement('div');
          novo.className = 'entry ok';
          novo.innerHTML = `
            <div class="d-flex justify-content-between align-items-start">
              <div>
                <div class="fw-bold text-success mb-1">${nome}</div>
                <div class="mt-1 ps-1">
                  <div><i class="bi bi-file-earmark-excel text-success me-1"></i>
                    <small><strong>Excel:</strong> ${nome} — ${formatCurrency(valor)} • ${hora}</small>
                  </div>
                  <div><i class="bi bi-file-earmark-pdf text-danger me-1"></i>
                    <small><strong>PDF:</strong> <em>confirmado manualmente</em></small>
                  </div>
                </div>
              </div>
              <button class="btn btn-sm btn-outline-danger desmarcar-conferido"
                      data-agente="${agente}" data-nome="${nome}"
                      data-valor="${valor}" data-hora="${hora}">
                <i class="bi bi-x-circle"></i>
              </button>
            </div>`;
          confContainer.insertAdjacentElement('afterend', novo);
          atualizarContadores(meta, confTitulo, faltandoTitulo, confCount + 1, faltandoCount - 1, valor, totalValor);
          novo.querySelector('.desmarcar-conferido').addEventListener('click', () =>
            moverItem(novo.querySelector('.desmarcar-conferido'), 'conferido', 'faltando'));
        }
      } else {
        // ➖ Volta para faltando PDF
        const faltContainer = document.querySelector(`#${agenteId} .fw-bold.text-warning`);
        if (faltContainer) {
          const novo = document.createElement('div');
          novo.className = 'entry warn';
          novo.innerHTML = `
            <div class="d-flex justify-content-between align-items-start">
              <div>
                <div class="fw-bold text-warning mb-1">${nome}</div>
                <div class="mt-1 ps-1">
                  <div><i class="bi bi-file-earmark-excel text-success me-1"></i>
                    <small><strong>Excel:</strong> ${nome} — ${formatCurrency(valor)} • ${hora}</small>
                  </div>
                  <div><i class="bi bi-file-earmark-pdf text-danger me-1"></i>
                    <small><strong>PDF:</strong> <em>não encontrado</em></small>
                  </div>
                </div>
              </div>
              <button class="btn btn-sm btn-outline-success marcar-conferido"
                      data-agente="${agente}" data-nome="${nome}"
                      data-valor="${valor}" data-hora="${hora}">
                <i class="bi bi-check-circle"></i>
              </button>
            </div>`;
          faltContainer.insertAdjacentElement('afterend', novo);
          atualizarContadores(meta, confTitulo, faltandoTitulo, confCount - 1, faltandoCount + 1, -valor, totalValor);
          novo.querySelector('.marcar-conferido').addEventListener('click', () =>
            moverItem(novo.querySelector('.marcar-conferido'), 'faltando', 'conferido'));
        }
      }
      recalcAndRenderAgent(agenteId);
    }


    // 🧮 Atualiza contadores e total dinamicamente
    function atualizarContadores(meta, confTitulo, faltandoTitulo, confCount, faltandoCount, valorDelta, totalAtual) {
      // Atualiza linha meta
      const metaText = meta?.textContent;
      let faltaExcel = 0;
      if (metaText) {
        const match = metaText.match(/Conferidos:\s*(\d+)\s*•\s*Falta PDF:\s*(\d+)\s*•\s*Falta Excel:\s*(\d+)/);
        if (match) {
          faltaExcel = parseInt(match[3]);
          meta.textContent = `Conferidos: ${Math.max(0, confCount)} • Falta PDF: ${Math.max(0, faltandoCount)} • Falta Excel: ${faltaExcel}`;
        }
      }

      // Atualiza total e contador de conferidos
      const totalSpan = confTitulo.querySelector('.total-conferidos');
      const novoTotal = Math.max(0, totalAtual + valorDelta);
      totalSpan.textContent = formatCurrency(novoTotal);
      confTitulo.innerHTML = `✅ Conferidos (${Math.max(0, confCount)}) — Total: <span class='total-conferidos'>${formatCurrency(novoTotal)}</span>`;

      // Atualiza contador de faltando no PDF
      if (faltandoTitulo) {
        faltandoTitulo.innerHTML = `⚠️ Faltando no PDF (${Math.max(0, faltandoCount)})`;
      }

      // 🔵 Atualiza círculo de porcentagem
      try {
        // usa o ID do agente (vem do meta → sobe pro .agent-content → pega o id)
        const agentContent = meta.closest('.agent-content');
        if (!agentContent) return;

        const agenteId = agentContent.id; // ex: GILMARIO_LIMA
        const agentCard = document.querySelector(`#${agenteId}`).closest('.agent-card');
        if (!agentCard) return;

        const total = Math.max(1, confCount + faltandoCount + faltaExcel);
        const perc = Math.round((confCount / total) * 100);

        const circles = agentCard.querySelectorAll('.progress-ring circle');
        const circle = circles[circles.length - 1]; // o círculo ativo (segundo)
        const inner = agentCard.querySelector('.circle-inner');

        if (circle && inner) {
          const r = 16;
          const circ = 2 * Math.PI * r;
          const offset = ((1 - perc / 100) * circ).toFixed(2);
          circle.style.transition = 'stroke-dashoffset 0.3s ease';
          circle.setAttribute('stroke-dashoffset', offset);
          circle.setAttribute('stroke', perc === 100 ? '#16a34a' : '#0a66c2');
          inner.textContent = `${perc}%`;
        }
      } catch (err) {
        console.warn('Erro ao atualizar círculo:', err);
      }
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
});

document.getElementById('btnExport').addEventListener('click', () => {
  const resultado = document.getElementById('resultado');
  if (!resultado.innerHTML) return alert('Nada para exportar');

  const totalC = document.getElementById('totalConferidos').textContent;
  const totalP = document.getElementById('totalFaltaPdf').textContent;
  const totalE = document.getElementById('totalFaltaExcel').textContent;
  const hoje = new Date();
  const dataStr = hoje.toLocaleDateString('pt-BR');
  const nomeArquivo = `ConferenciaCaixa_${bancoDetectado || 'DESCONHECIDO'}_${hoje.toISOString().split('T')[0]}.pdf`;

  const cabecalho = `
    <div style='text-align:center;margin-bottom:20px;'>
      <h2 style='color:#0a66c2;margin-bottom:4px;'>📊 Conferência de Caixa</h2>
      <p style='margin:0;font-size:13px;color:#444;'>Banco: <strong>${bancoDetectado}</strong> • Data: <strong>${dataStr}</strong></p>
      <p style='margin:4px 0;font-size:13px;color:#555;'>Conferidos: ${totalC} • Falta PDF: ${totalP} • Falta Excel: ${totalE}</p>
      <hr style='border:none;border-top:1px solid #ccc;margin:10px 0;'>
    </div>`;

  const conteudoPDF = cabecalho + resultado.innerHTML;
  const opt = { margin: 0.5, filename: nomeArquivo, html2canvas: { scale: 2 }, jsPDF: { unit: 'in', format: 'a4', orientation: 'portrait' } };
  html2pdf().set(opt).from(conteudoPDF).save();
});
