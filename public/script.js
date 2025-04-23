document.addEventListener('DOMContentLoaded', () => {
  const senha = prompt('Por favor, insira a senha para acessar o sistema:');
  if (senha !== 'planilha') {
	document.body.innerHTML = `
	  <div style="display:flex;align-items:center;justify-content:center;height:100vh;font-family:Arial, sans-serif;">
		<h2>Acesso negado.</h2>
	  </div>
	`;
	return;
  }

  const fileInput = document.getElementById('file');
  const titleInput = document.getElementById('title');
  const warnEl = document.getElementById('warning');
  const processBtn = document.getElementById('process');

  processBtn.addEventListener('click', async () => {
	if (!fileInput.files.length) {
	  alert('Selecione um arquivo .xlsx');
	  return;
	}

	const file = fileInput.files[0];
	const title = titleInput.value.trim() || 'download';

	const form = new FormData();
	form.append('file', file);
	form.append('title', title);

	const resp = await fetch('/upload', { method: 'POST', body: form });
	if (!resp.ok) {
	  alert('Erro ao processar o arquivo');
	  return;
	}

	const failed = resp.headers.get('X-CNPJ-Failure') === '1';
	warnEl.style.display = failed ? 'block' : 'none';

	const blob = await resp.blob();
	const url = URL.createObjectURL(blob);
	const a = document.createElement('a');
	a.href = url;
	a.download = `${title}.xlsx`;
	document.body.appendChild(a);
	a.click();
	a.remove();
	URL.revokeObjectURL(url);
  });
});