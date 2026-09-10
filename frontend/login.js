document.getElementById('loginForm').addEventListener('submit', async (event) => {
  event.preventDefault();
  const button = document.getElementById('btnEntrar');
  const error = document.getElementById('loginError');
  error.hidden = true;
  button.disabled = true;
  button.textContent = 'Entrando…';
  try {
    const response = await fetch('/auth/login', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ usuario: document.getElementById('usuario').value.trim(), senha: document.getElementById('senha').value })
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.erro || 'Não foi possível entrar. Tente novamente.');
    location.replace('/');
  } catch (e) {
    error.textContent = e instanceof SyntaxError ? 'Não foi possível entrar. Tente novamente.' : e.message;
    error.hidden = false;
    document.getElementById('senha').value = '';
  } finally {
    button.disabled = false;
    button.textContent = 'Entrar';
  }
});
