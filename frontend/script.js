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
    const resp = await fetch('http://127.0.0.1:8000/conferir_caixa', { method: 'POST', body: fd });
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
              <span class='agent-name'><i class='bi bi-person-circle'></i> ${agente}</span><br>
              <span class='agent-meta'>Conferidos: ${d.conferidos.length} • Falta PDF: ${d.faltando_pdf.length} • Falta Excel: ${d.faltando_excel.length}</span>
            </div>
            ${circle}
          </div>
          <div class='agent-content' id='${id}'>
            <div class='mt-3'>
              <div class='fw-bold text-success mb-2'>✅ Conferidos (${d.conferidos.length})</div>
              ${d.conferidos.map(x => `
                <div class='entry ok'>
                  <div class="fw-bold text-success mb-1">${x.nome_excel || x.nome}</div>
                  <div class="mt-1 ps-1">
                    <div><i class="bi bi-file-earmark-excel text-success me-1"></i>
                      <small><strong>Excel:</strong> ${x.nome_excel || '-'} — ${formatCurrency(x.valor_excel)} • ${x.hora_excel || '(sem hora)'}</small>
                    </div>
                    <div><i class="bi bi-file-earmark-pdf text-danger me-1"></i>
                      <small><strong>PDF:</strong> ${x.nome_pdf || '-'} — ${formatCurrency(x.valor_pdf)} • ${x.hora_pdf || '(sem hora)'}</small>
                    </div>
                  </div>
                </div>`).join('')}

              <div class='fw-bold text-warning mt-3 mb-2'>⚠️ Faltando no PDF (${d.faltando_pdf.length})</div>
              ${d.faltando_pdf.map(x => `
                <div class='entry warn'>
                  <strong>${x.nome}</strong>
                  <div class="mt-1">
                    <div><i class="bi bi-file-earmark-excel text-success"></i>
                      <small><strong>Excel:</strong> ${formatCurrency(x.valor_excel ?? x.valor)} • ${x.hora || '(sem hora)'}</small>
                    </div>
                    <div><i class="bi bi-file-earmark-pdf text-danger"></i>
                      <small><strong>PDF:</strong> <em>não encontrado</em></small>
                    </div>
                  </div>
                  <div class="text-muted mt-1"><small>💬 ${x.motivo || 'Sem motivo registrado.'}</small></div>
                </div>`).join('')}

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
