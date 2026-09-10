// O Excel gerado é mantido apenas em memória e enviado à conferência existente.
let planilhaSheets = null;
let sheetsBusy = false;
const fonteDados = document.getElementById('fonteDados');
const sheetsStatus = document.getElementById('sheetsStatus');
const dateParts = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date());
const datePart = type => dateParts.find(part => part.type === type).value;
document.getElementById('sheetsData').value = `${datePart('year')}-${datePart('month')}-${datePart('day')}`;

function resetSheets() {
  planilhaSheets = null;
  sheetsStatus.textContent = 'Escolha a data para buscar os lançamentos.';
  sheetsStatus.classList.remove('text-danger', 'text-success');
}

function controlFiles() {
  return fonteDados.value === 'sheets' ? (planilhaSheets ? [planilhaSheets] : []) : [...document.getElementById('excelFile').files];
}

fonteDados.addEventListener('change', () => {
  document.getElementById('sheetsControls').hidden = fonteDados.value !== 'sheets';
  document.getElementById('excelControls').hidden = fonteDados.value !== 'excel';
  resetSheets();
});
document.getElementById('sheetsData').addEventListener('change', resetSheets);
document.getElementById('sheetsTipo').addEventListener('change', () => {
  resetSheets();
  document.getElementById('sheetsGroupHint').textContent = document.getElementById('sheetsTipo').value === 'geral'
    ? 'Geral inclui todos os agentes da base.'
    : 'Agência: David Elias, Ermesson Lima, Livia Oliveira, Mathias Lima, Noe Lemos e Vitoria Regia.';
});

document.getElementById('btnImportarSheets').addEventListener('click', async () => {
  if (sheetsBusy || !document.getElementById('sheetsData').reportValidity()) return;
  resetSheets();
  const data = document.getElementById('sheetsData').value;
  const tipo = document.getElementById('sheetsTipo').value;
  const ids = ['fonteDados', 'sheetsData', 'sheetsTipo', 'btnImportarSheets', 'btnConferir', 'btnLimpar'];
  sheetsBusy = true;
  ids.forEach(id => document.getElementById(id).disabled = true);
  sheetsStatus.textContent = 'Buscando os lançamentos no Google Sheets…';
  const abort = new AbortController();
  const timer = setTimeout(() => abort.abort(), 120000);
  try {
    const response = await fetch('/api/sheets/importar', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data, tipo }), signal: abort.signal
    });
    if (response.status === 401) { location.replace('/login'); return; }
    if (!response.ok) {
      const problem = await response.json().catch(() => ({}));
      throw new Error(problem.erro || 'Não foi possível carregar os dados. Tente novamente.');
    }
    planilhaSheets = new File([await response.blob()], `TABELA_${tipo.toUpperCase()}_${data}.xlsx`, {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const formattedDate = data.split('-').reverse().join('/');
    sheetsStatus.textContent = `${response.headers.get('X-Lancamentos')} lançamentos de ${response.headers.get('X-Agentes')} agentes carregados · ${formattedDate}. Envie o PDF e inicie a conferência.`;
    sheetsStatus.classList.add('text-success');
  } catch (error) {
    resetSheets();
    sheetsStatus.textContent = error.name === 'AbortError' ? 'A consulta demorou demais. Tente novamente.' : error.message;
    sheetsStatus.classList.add('text-danger');
  } finally {
    clearTimeout(timer);
    sheetsBusy = false;
    ids.forEach(id => document.getElementById(id).disabled = false);
  }
});

document.getElementById('btnSair').addEventListener('click', async () => {
  const button = document.getElementById('btnSair');
  button.disabled = true;
  try {
    const response = await fetch('/auth/logout', { method: 'POST' });
    if (!response.ok && response.status !== 401) throw new Error('Não foi possível sair. Tente novamente.');
    planilhaSheets = null;
    location.replace('/login');
  } catch (error) {
    alert(error.message);
    button.disabled = false;
  }
});
