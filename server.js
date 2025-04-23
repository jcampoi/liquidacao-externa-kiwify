const express = require('express');
const multer  = require('multer');
const ExcelJS = require('exceljs');
const axios = require('axios');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

const CNPJ_API_URL = 'https://www.receitaws.com.br/v1/cnpj/';

app.use(express.static(path.join(__dirname, 'public')));

async function fetchRazaoSocial(cnpj) {
  try {
	const resp = await axios.get(`${CNPJ_API_URL}${cnpj}`);
	if (resp.data && resp.data.nome) return resp.data.nome;
	return '';
  } catch (err) {
	console.error(`Erro ao buscar CNPJ ${cnpj}:`, err.message);
	return 'Não foi possível procurar';
  }
}

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
	const title = (req.body.title || 'download').trim();
	const buffer = req.file.buffer;

	const wb = new ExcelJS.Workbook();
	await wb.xlsx.load(buffer);
	const ws = wb.worksheets[0];

	const desired = [
	  'Identificador do Recebível',
	  'total efeitos de contrato capturado',
	  'Usuário recebbedor final',
	  'Tipo de efeito',
	  'Id do contrato',
	  'CNPJ da empresa que gerou contrato'
	];
	const renameMap = {
	  'Identificador do Recebível': 'Identificador do Recebível',
	  'total efeitos de contrato capturado': 'Total efeitos de contrato',
	  'Usuário recebbedor final': 'Usuário',
	  'Tipo de efeito': 'Tipo de efeito',
	  'Id do contrato': 'Id do contrato',
	  'CNPJ da empresa que gerou contrato': 'CNPJ'
	};

	const header = ws.getRow(1);
	const cols = [];
	header.eachCell((cell, colNumber) => {
	  if (desired.includes(cell.value)) cols.push(colNumber);
	});

	const cnpjOrigIndex = cols[cols.length - 1];

	const newWb = new ExcelJS.Workbook();
	const newWs = newWb.addWorksheet('Sheet1');

	const headerNames = cols.map(ci => renameMap[header.getCell(ci).value]);
	const insertIndex = headerNames.indexOf('Id do contrato') + 1;
	headerNames.splice(insertIndex, 0, 'Empresa que gerou o contrato');

	const newHeader = newWs.addRow(headerNames);
	newHeader.eachCell(cell => {
	  cell.font = { bold: true };
	  cell.alignment = { horizontal: 'center' };
	  cell.fill = {
		type: 'pattern',
		pattern: 'solid',
		fgColor: { argb: 'FFC2D79C' }
	  };
	});

	const razaoCache = {};
	let anyFailure = false;

	const dataRows = [];
	ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
	  if (rowNumber > 1) dataRows.push(row);
	});

	for (const row of dataRows) {
	  const vals = cols.map(ci => row.getCell(ci).value);
	  const cnpj = row.getCell(cnpjOrigIndex).value;

	  let razao;
	  if (razaoCache[cnpj] !== undefined) {
		razao = razaoCache[cnpj];
	  } else {
		razao = await fetchRazaoSocial(cnpj);
		razaoCache[cnpj] = razao;
	  }
	  if (razao === '') anyFailure = true;

	  vals.splice(insertIndex, 0, razao);
	  const newRow = newWs.addRow(vals);

	  cols.forEach((ci, i) => {
		const origCell = row.getCell(ci);
		const targetCol = i < insertIndex ? i + 1 : i + 2;
		if (origCell.numFmt) newRow.getCell(targetCol).numFmt = origCell.numFmt;
	  });
	}

	res.setHeader('X-CNPJ-Failure', anyFailure ? '1' : '0');

	res.setHeader('Content-Disposition', `attachment; filename="${title}.xlsx"`);
	res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	await newWb.xlsx.write(res);
	res.end();

  } catch (err) {
	console.error(err);
	res.status(500).send('Erro ao processar o arquivo');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Rodando em http://localhost:${PORT}`));