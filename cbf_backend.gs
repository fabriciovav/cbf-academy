// ============================================================
// CBF ACADEMY — Backend Google Apps Script
// Recebe dados do formulário e serve ao painel admin
//
// COMO CONFIGURAR:
// 1. Acesse script.google.com → Novo projeto
// 2. Cole este código (apague o myFunction)
// 3. Clique em "Implantar" → "Nova implantação"
// 4. Tipo: "Aplicativo da Web"
// 5. Executar como: "Eu"
// 6. Quem tem acesso: "Qualquer pessoa"
// 7. Clique em "Implantar" e copie a URL gerada
// 8. Cole essa URL no cbf_formulario.html e cbf_admin.html
//    onde está escrito: COLE_AQUI_A_URL_DO_WEB_APP
// ============================================================

const SHEET_NAME = 'Avaliações';

// Recebe dados do formulário (POST)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    salvarResposta(data);
    return ContentService
      .createTextOutput(JSON.stringify({status: 'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', msg: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Serve dados ao painel admin (GET)
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getData') {
    const dados = lerTodosOsDados();
    return ContentService
      .createTextOutput(JSON.stringify(dados))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput('CBF Academy API funcionando.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// Salva uma linha na planilha
function salvarResposta(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Criar aba se não existir
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Cabeçalhos
    const headers = [
      'Timestamp', 'Treinador', 'Data', 'Licença', 'Módulo', 'Avaliador',
      'Local', 'Categoria', 'Nº Atletas',
      'D1.1','D1.2','D1.3','D1.4',
      'D2.1','D2.2','D2.3','D2.4','D2.5',
      'D3.1','D3.2','D3.3','D3.4','D3.5',
      'D4.1','D4.2','D4.3',
      'D5.1','D5.2','D5.3',
      'D1.1_obs','D1.2_obs','D1.3_obs','D1.4_obs',
      'D2.1_obs','D2.2_obs','D2.3_obs','D2.4_obs','D2.5_obs',
      'D3.1_obs','D3.2_obs','D3.3_obs','D3.4_obs','D3.5_obs',
      'D4.1_obs','D4.2_obs','D4.3_obs',
      'D5.1_obs','D5.2_obs','D5.3_obs',
      'Pontos Fortes','Áreas de Desenvolvimento','Recomendações'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#00703C')
      .setFontColor('white')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Montar linha
  const row = [
    data.timestamp || new Date().toISOString(),
    data.treinador || '',
    data.data || '',
    data.licenca || '',
    data.modulo || '',
    data.avaliador || '',
    data.local || '',
    data.categoria || '',
    data.atletas || '',
    data['D1.1'] !== undefined ? data['D1.1'] : '',
    data['D1.2'] !== undefined ? data['D1.2'] : '',
    data['D1.3'] !== undefined ? data['D1.3'] : '',
    data['D1.4'] !== undefined ? data['D1.4'] : '',
    data['D2.1'] !== undefined ? data['D2.1'] : '',
    data['D2.2'] !== undefined ? data['D2.2'] : '',
    data['D2.3'] !== undefined ? data['D2.3'] : '',
    data['D2.4'] !== undefined ? data['D2.4'] : '',
    data['D2.5'] !== undefined ? data['D2.5'] : '',
    data['D3.1'] !== undefined ? data['D3.1'] : '',
    data['D3.2'] !== undefined ? data['D3.2'] : '',
    data['D3.3'] !== undefined ? data['D3.3'] : '',
    data['D3.4'] !== undefined ? data['D3.4'] : '',
    data['D3.5'] !== undefined ? data['D3.5'] : '',
    data['D4.1'] !== undefined ? data['D4.1'] : '',
    data['D4.2'] !== undefined ? data['D4.2'] : '',
    data['D4.3'] !== undefined ? data['D4.3'] : '',
    data['D5.1'] !== undefined ? data['D5.1'] : '',
    data['D5.2'] !== undefined ? data['D5.2'] : '',
    data['D5.3'] !== undefined ? data['D5.3'] : '',
    data['D1.1_obs'] || '', data['D1.2_obs'] || '',
    data['D1.3_obs'] || '', data['D1.4_obs'] || '',
    data['D2.1_obs'] || '', data['D2.2_obs'] || '',
    data['D2.3_obs'] || '', data['D2.4_obs'] || '',
    data['D2.5_obs'] || '', data['D3.1_obs'] || '',
    data['D3.2_obs'] || '', data['D3.3_obs'] || '',
    data['D3.4_obs'] || '', data['D3.5_obs'] || '',
    data['D4.1_obs'] || '', data['D4.2_obs'] || '',
    data['D4.3_obs'] || '', data['D5.1_obs'] || '',
    data['D5.2_obs'] || '', data['D5.3_obs'] || '',
    data.pontos_fortes || '',
    data.areas_desenvolvimento || '',
    data.recomendacoes || '',
  ];

  sheet.appendRow(row);
}

// Lê todos os dados e retorna como JSON
function lerTodosOsDados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  const rows = values.slice(1);

  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i] !== undefined ? row[i] : '';
    });
    // Mapear para nomes esperados pelo admin
    return {
      timestamp: obj['Timestamp'],
      treinador: obj['Treinador'],
      data: obj['Data'] ? String(obj['Data']).substring(0, 10) : '',
      licenca: obj['Licença'],
      modulo: obj['Módulo'],
      avaliador: obj['Avaliador'],
      local: obj['Local'],
      categoria: obj['Categoria'],
      atletas: obj['Nº Atletas'],
      'D1.1': obj['D1.1'], 'D1.2': obj['D1.2'],
      'D1.3': obj['D1.3'], 'D1.4': obj['D1.4'],
      'D2.1': obj['D2.1'], 'D2.2': obj['D2.2'],
      'D2.3': obj['D2.3'], 'D2.4': obj['D2.4'],
      'D2.5': obj['D2.5'],
      'D3.1': obj['D3.1'], 'D3.2': obj['D3.2'],
      'D3.3': obj['D3.3'], 'D3.4': obj['D3.4'],
      'D3.5': obj['D3.5'],
      'D4.1': obj['D4.1'], 'D4.2': obj['D4.2'],
      'D4.3': obj['D4.3'],
      'D5.1': obj['D5.1'], 'D5.2': obj['D5.2'],
      'D5.3': obj['D5.3'],
      pontos_fortes: obj['Pontos Fortes'],
      areas_desenvolvimento: obj['Áreas de Desenvolvimento'],
      recomendacoes: obj['Recomendações'],
    };
  });
}
