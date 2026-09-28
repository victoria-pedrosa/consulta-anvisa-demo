/**
 * MOTOR DE VALIDACAO E EVIDENCIAS - LC 214/2025
 * Comprovacao de enquadramento de produtos no tratamento tributario diferenciado.
 *
 * ETAPA A - ANVISA:   o produto esta regularizado?
 * ETAPA B - LC 214:   o produto esta enquadrado na hipotese que gera o beneficio?
 * As duas respostas sao independentes e o resultado final so e favoravel quando ambas fecham.
 *
 * Base legal: LC 214/2025, art. 131, caput e §1o (dispositivos do Anexo IV regularizados
 * na Anvisa); art. 131, §2o (revisao dos anexos a cada 120 dias por ato conjunto do
 * Ministro da Fazenda e do Comite Gestor do IBS, ouvido o Ministerio da Saude);
 * art. 133 (medicamentos registrados na Anvisa); art. 144 (aplica o §2o do art. 131).
 *
 * Todo o parsing de CSV/XML e o cruzamento rodam no navegador (Index.html).
 * Este arquivo serve a pagina, le e grava abas, chama a API Gemini e emite o relatorio.
 */

var ABA_BASE      = 'BASE_ANVISA';
var ABA_LC214     = 'BASE_LC214';
var ABA_MANUAL    = 'CADASTRO_MANUAL';
var ABA_RESULTADO = 'PROVA_ANVISA';
var ABA_DECISAO   = 'MOTOR_DECISAO';
var ABA_LOG       = 'LOG_EXECUCAO';

var GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta';

/* ============================================================
 * CHAVE DA API GEMINI - cole entre as aspas.
 * Fica dentro do script, visivel para quem tiver acesso de edicao ao projeto.
 * Nao compartilhe o projeto do Apps Script com quem nao deve ver a chave.
 * ============================================================ */
var CHAVE_GEMINI = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

/* Modelo padrao. O botao "Conferir modelos" lista o que a chave realmente acessa
   e a chamada cai automaticamente para um modelo valido se este nao existir. */
var MODELO_PADRAO = 'gemini-3.7-flash';

/* ===================== pagina ===================== */

/* Menu na planilha, para quem nunca vai abrir o editor de script. */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Prova ANVISA')
      .addItem('Abrir a ferramenta', 'mostrarLateral')
      .addSeparator()
      .addItem('Instalar/atualizar a base da LC 214', 'instalarBaseLC214')
      .addItem('Criar as abas de apoio', 'prepararAbas')
      .addToUi();
  } catch (e) {}
}

function mostrarLateral() {
  var html = HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Prova ANVISA - LC 214/2025').setWidth(480);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Grava a base da LC 214 que vem embutida no projeto (BaseLC214.gs) na aba
 * BASE_LC214. Nao depende de download, de internet nem de colar planilha.
 */
function instalarBaseLC214() {
  try {
    var m = matrizBaseLC214();
    var aba = gravarMatriz_(ABA_LC214, m);
    aba.getRange(2, 1, m.length - 1, 1).setNumberFormat('@');
    aba.setColumnWidth(1, 90); aba.setColumnWidth(8, 520);
    PropertiesService.getDocumentProperties().setProperty('versaoBaseLC214', VERSAO_BASE_LC214);
    var msg = m.length - 1 + ' linhas gravadas na aba ' + ABA_LC214 + '.';
    try { SpreadsheetApp.getUi().alert('Base da LC 214 instalada', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
    return { ok: true, linhas: m.length - 1, versao: VERSAO_BASE_LC214 };
  } catch (e) {
    return { ok: false, erro: String(e) };
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Motor de Evidencias - LC 214/2025')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

function ss_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Script nao esta vinculado a uma planilha.');
  return ss;
}

function aba_(nome, cabecalho) {
  var ss = ss_();
  var aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    if (cabecalho) {
      aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho])
         .setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');
      aba.setFrozenRows(1);
    }
  }
  return aba;
}

/* ===================== abas de apoio ===================== */

function prepararAbas() {
  try {
    aba_(ABA_MANUAL, ['Codigo/GTIN', 'Produto', 'Tipo', 'Numero ANVISA', 'Validade',
                      'Conferido por', 'Data', 'Observacao'])
      .getRange('C2:C').setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['REGISTRO', 'CADASTRO', 'NOTIFICACAO', 'ISENTO', 'NAO REGULARIZADO'], true)
          .build());
    aba_(ABA_LC214, ['NCM', 'Item', 'Anexo', 'Hipotese tributaria', 'Beneficio', 'Exige Anvisa (S/N)']);
    aba_(ABA_LOG, ['Data/Hora', 'Usuario', 'Arquivo produtos', 'XMLs', 'Produtos',
                   'Base ANVISA', 'Data da base', 'Hash da base', 'IA', 'Modelo',
                   'Nivel A', 'Nivel B', 'Nivel C', 'Nivel D', 'Resultado']);
    return { ok: true };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

function lerAba(nome) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { ok: true, linhas: [] };
    var aba = ss.getSheetByName(nome);
    if (!aba || aba.getLastRow() < 2) return { ok: true, linhas: [] };
    return { ok: true, linhas: aba.getDataRange().getValues() };
  } catch (e) { return { ok: false, erro: String(e), linhas: [] }; }
}

function lerBaseAnvisa()     { return lerAba(ABA_BASE); }
function lerCadastroManual() { return lerAba(ABA_MANUAL); }
/**
 * A aba BASE_LC214 tem prioridade - e onde a equipe pode ajustar algum
 * enquadramento. Vazia, o app usa a base embutida no proprio projeto.
 */
function lerBaseLC214() {
  var r = lerAba(ABA_LC214);
  if (r.ok && r.linhas && r.linhas.length > 1) {
    r.origem = 'aba ' + ABA_LC214;
    return r;
  }
  return { ok: true, linhas: matrizBaseLC214(), origem: 'base embutida no script', versao: VERSAO_BASE_LC214 };
}

function anexarCadastroManual(linhas) {
  try {
    if (!linhas || !linhas.length) return { ok: true, linhas: 0 };
    prepararAbas();
    var aba = ss_().getSheetByName(ABA_MANUAL);
    aba.getRange(aba.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
    return { ok: true, linhas: linhas.length };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/* ===================== gravacao de resultado ===================== */

function gravarMatriz_(nome, matriz) {
  var aba = ss_().getSheetByName(nome);
  if (aba) { aba.clear(); } else { aba = ss_().insertSheet(nome); }
  var passo = 2000;
  for (var i = 0; i < matriz.length; i += passo) {
    var bloco = matriz.slice(i, i + passo);
    aba.getRange(i + 1, 1, bloco.length, matriz[0].length).setValues(bloco);
  }
  aba.getRange(1, 1, 1, matriz[0].length).setFontWeight('bold')
     .setBackground('#1f3864').setFontColor('#ffffff');
  aba.setFrozenRows(1);
  return aba;
}

function salvarResultado(matriz, matrizDecisao) {
  try {
    if (!matriz || matriz.length < 2) throw new Error('Nada para gravar.');
    gravarMatriz_(ABA_RESULTADO, matriz);
    if (matrizDecisao && matrizDecisao.length > 1) gravarMatriz_(ABA_DECISAO, matrizDecisao);
    return { ok: true, linhas: matriz.length - 1, abas: [ABA_RESULTADO, ABA_DECISAO] };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

function gravarBaseAnvisa(matriz) {
  try { gravarMatriz_(ABA_BASE, matriz); return { ok: true, linhas: matriz.length }; }
  catch (e) { return { ok: false, erro: String(e) }; }
}

function registrarLog(d) {
  try {
    prepararAbas();
    var aba = ss_().getSheetByName(ABA_LOG);
    aba.appendRow([
      new Date(), (Session.getActiveUser() && Session.getActiveUser().getEmail()) || 'n/d',
      d.arquivoProdutos || '', d.qtdXml || 0, d.produtos || 0,
      d.arquivoBase || '', d.dataBase || '', d.hashBase || '',
      d.usouIA ? 'Sim' : 'Nao', d.modelo || '',
      d.a || 0, d.b || 0, d.c || 0, d.dd || 0, d.resultado || 'Concluido'
    ]);
    return { ok: true };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/* ===================== base da ANVISA: download, hash e validade ===================== */

function hash_(texto) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  return b.slice(0, 8).map(function (x) { return ('0' + (x & 0xFF).toString(16)).slice(-2); }).join('');
}

function baixarCsvAnvisa(url) {
  try {
    if (!url || url.indexOf('https://dados.anvisa.gov.br/') !== 0) {
      throw new Error('URL precisa comecar com https://dados.anvisa.gov.br/');
    }
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
    var txt = r.getContentText('UTF-8');
    return { ok: true, texto: txt, arquivo: url.split('/').pop(), hash: hash_(txt) };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

function listarArquivosAnvisa(caminho) {
  try {
    var url = caminho || 'https://dados.anvisa.gov.br/dados/';
    if (url.indexOf('https://dados.anvisa.gov.br/') !== 0) throw new Error('Caminho fora do dominio de dados abertos.');
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
    var html = r.getContentText(), achados = {}, m, re = /href="([^"?#]+)"/gi;
    while ((m = re.exec(html)) !== null) {
      var h = m[1];
      if (h.indexOf('..') === 0 || h === '/' || h.indexOf('?') === 0) continue;
      var abs = h.indexOf('http') === 0 ? h : url.replace(/\/+$/, '/') + h.replace(/^\/+/, '');
      if (/\.(csv|zip|txt)$/i.test(abs) || /\/$/.test(abs)) achados[abs] = true;
    }
    return { ok: true, url: url, arquivos: Object.keys(achados).sort() };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/**
 * Identificacao da base usada no estudo: nome do arquivo, data e hash.
 * O hash permite provar depois QUAL arquivo exatamente sustentou a conclusao.
 */
function registrarCargaBase(info) {
  PropertiesService.getDocumentProperties().setProperty('baseAnvisa', JSON.stringify({
    data: new Date().toISOString(), linhas: info.linhas || 0,
    arquivo: info.arquivo || 'upload manual', hash: info.hash || ''
  }));
  return { ok: true };
}

/**
 * Idade da base da Anvisa (Etapa A). A Anvisa atualiza a base continuamente e nao ha
 * prazo legal de validade - o semaforo aqui e operacional, so indica quao velha esta a copia.
 */
/**
 * Versao dos Anexos da LC 214 efetivamente usada na Etapa B.
 * O ciclo de 120 dias do art. 131, §2o e do ANEXO, nao da base da Anvisa - sao
 * dois relogios diferentes e confundi-los invalida a rastreabilidade do estudo.
 */
function registrarVersaoAnexos(info) {
  PropertiesService.getDocumentProperties().setProperty('anexosLC214', JSON.stringify({
    data: info.data || '', ato: info.ato || '', registradoEm: new Date().toISOString()
  }));
  return { ok: true };
}

function statusAnexos() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty('anexosLC214');
    if (!raw) return { ok: true, tem: false };
    var d = JSON.parse(raw);
    if (!d.data) return { ok: true, tem: false };
    var dias = Math.floor((new Date() - new Date(d.data)) / 86400000);
    return { ok: true, tem: true, dias: dias, ato: d.ato,
             faixa: dias <= 90 ? 'ATUAL' : (dias <= 120 ? 'ATENCAO' : 'VERIFICAR'),
             dataTexto: Utilities.formatDate(new Date(d.data), Session.getScriptTimeZone(), 'dd/MM/yyyy') };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

function statusBase() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty('baseAnvisa');
    if (!raw) return { ok: true, temBase: false };
    var d = JSON.parse(raw);
    var dias = Math.floor((new Date() - new Date(d.data)) / 86400000);
    return {
      ok: true, temBase: true, dias: dias, linhas: d.linhas,
      arquivo: d.arquivo, hash: d.hash,
      faixa: dias <= 90 ? 'ATUAL' : (dias <= 180 ? 'ATENCAO' : 'ANTIGA'),
      dataTexto: Utilities.formatDate(new Date(d.data), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      dataIso: d.data
    };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/* ===================== motor de IA ===================== */

function temChaveIA() {
  return { ok: true, tem: !!chaveValida_(), modelo: MODELO_PADRAO };
}

function chaveValida_() {
  var k = String(CHAVE_GEMINI || '').trim();
  return (k && k.indexOf('COLE_AQUI') < 0 && k.length >= 20) ? k : '';
}

function chave_() {
  var k = chaveValida_();
  if (!k) throw new Error('Preencha a constante CHAVE_GEMINI no arquivo Codigo.gs.');
  return k;
}

/** Primeiro modelo valido da conta, quando o configurado nao existir mais. */
function modeloAlternativo_() {
  var r = listarModelosIA();
  if (!r.ok || !r.modelos.length) return '';
  var flash = r.modelos.filter(function (m) { return m.indexOf('flash') > 0 && m.indexOf('lite') < 0; });
  return (flash.length ? flash : r.modelos)[0];
}

function listarModelosIA() {
  try {
    var r = UrlFetchApp.fetch(GEMINI_BASE + '/models?key=' + encodeURIComponent(chave_()),
                              { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
    var ms = (JSON.parse(r.getContentText()).models || [])
      .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
      .map(function (m) { return m.name.replace('models/', ''); })
      .filter(function (n) { return n.indexOf('gemini') === 0; });
    return { ok: true, modelos: ms };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/**
 * A IA escolhe entre candidatos REAIS e precisa declarar quais atributos comparou.
 * Nome parecido, sozinho, nao autoriza conclusao - o cliente rebaixa a resposta que
 * confirmar menos de dois atributos alem da descricao.
 */
var _jaTrocouModelo = false;
function classificarComIA(lote, modelo) {
  try {
    modelo = modelo || MODELO_PADRAO;
    var instrucao =
      'Voce e analista de regularizacao sanitaria. Para cada item, decida qual dos candidatos ' +
      'e o MESMO produto. Compare, um a um: descricao/nome tecnico, fabricante ou detentor, ' +
      'modelo, apresentacao, volume ou tamanho, unidade de medida, GTIN e codigo. ' +
      'REGRAS RIGIDAS: ' +
      '(1) escolha apenas um id presente na lista de candidatos daquele item, ou null; ' +
      '(2) nunca invente numero de registro, fabricante ou dado que nao esteja nos candidatos; ' +
      '(3) semelhanca apenas de nome NAO e suficiente - se so a descricao bate e nenhum outro ' +
      'atributo pode ser confirmado, responda escolhido null; ' +
      '(4) atributos_compativeis deve listar SO os atributos que voce efetivamente confirmou ' +
      'com base nos dados recebidos, e atributos_divergentes os que nao batem; ' +
      '(5) confianca abaixo de 0.8 sempre que houver divergencia de volume, tamanho, ' +
      'apresentacao ou fabricante; ' +
      '(6) motivos: ate 4 frases curtas, cada uma citando o atributo comparado.';

    var corpo = {
      systemInstruction: { parts: [{ text: instrucao }] },
      contents: [{ role: 'user', parts: [{ text: JSON.stringify(lote) }] }],
      generationConfig: {
        temperature: 0,
        responseMimeType: 'application/json',
        responseSchema: {
          type: 'ARRAY',
          items: {
            type: 'OBJECT',
            properties: {
              i:                     { type: 'INTEGER' },
              escolhido:             { type: 'STRING' },
              confianca:             { type: 'NUMBER' },
              atributos_compativeis: { type: 'ARRAY', items: { type: 'STRING' } },
              atributos_divergentes: { type: 'ARRAY', items: { type: 'STRING' } },
              motivos:               { type: 'ARRAY', items: { type: 'STRING' } }
            },
            required: ['i', 'escolhido', 'confianca', 'atributos_compativeis', 'motivos']
          }
        }
      }
    };

    var r = UrlFetchApp.fetch(GEMINI_BASE + '/models/' + modelo + ':generateContent?key=' + encodeURIComponent(chave_()), {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(corpo), muteHttpExceptions: true
    });
    var cod = r.getResponseCode();
    if (cod === 429) return { ok: false, erro: 'Limite de requisicoes da API atingido.', retry: true };
    if (cod === 404 && !_jaTrocouModelo) {          /* modelo inexistente: troca uma vez */
      var alt = modeloAlternativo_();
      if (alt && alt !== modelo) {
        _jaTrocouModelo = true;
        var novo = classificarComIA(lote, alt);
        if (novo.ok) novo.modeloUsado = alt;
        return novo;
      }
    }
    if (cod !== 200) throw new Error('HTTP ' + cod + ' ' + r.getContentText().slice(0, 300));

    var j = JSON.parse(r.getContentText());
    var txt = j.candidates && j.candidates[0] && j.candidates[0].content &&
              j.candidates[0].content.parts && j.candidates[0].content.parts[0].text;
    if (!txt) throw new Error('Resposta vazia do modelo.');
    return { ok: true, itens: JSON.parse(txt) };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/* ===================== relatorio da prova ===================== */

/**
 * Gera um documento no Drive com a metodologia, a identificacao da base e os numeros.
 * O que e classificado como PROVAVEL nao entra como comprovacao - fica registrado no texto.
 */
function gerarRelatorio(d) {
  try {
    var nome = 'Prova de Regularizacao LC 214 - ' + (d.empresa || 'cliente') + ' - ' +
               Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var doc = DocumentApp.create(nome);
    var b = doc.getBody();

    b.appendParagraph('PROVA DE REGULARIZACAO E ENQUADRAMENTO - LC 214/2025')
     .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    b.appendParagraph('1. Identificacao').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    var ident = [
      ['Empresa', d.empresa || '-'],
      ['CNPJ', d.cnpj || '-'],
      ['Periodo analisado', d.periodo || '-'],
      ['Produtos analisados', String(d.produtos || 0)],
      ['Arquivo de produtos', d.arquivoProdutos || '-'],
      ['XMLs de NF-e lidos', String(d.qtdXml || 0)],
      ['Base ANVISA utilizada', d.arquivoBase || '-'],
      ['Data de download da base', d.dataBase || '-'],
      ['Versao dos Anexos da LC 214 (Etapa B)', d.versaoAnexos || 'nao informada'],
      ['Ato conjunto de referencia', d.atoAnexos || '-'],
      ['Identificador da base (SHA-256, 8 bytes)', d.hashBase || '-'],
      ['IA utilizada', d.usouIA ? ('Sim - ' + (d.modelo || '')) : 'Nao'],
      ['Confianca minima exigida da IA', d.usouIA ? String(d.confMin) : '-'],
      ['Emitido em', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')]
    ];
    b.appendTable(ident).setBorderWidth(0.5);

    b.appendParagraph('2. Base legal').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    b.appendParagraph(
      'LC 214/2025, art. 131, caput e §1o - reducao de 60% das aliquotas de IBS e CBS sobre os ' +
      'dispositivos medicos relacionados no Anexo IV, desde que regularizados perante a Anvisa. ' +
      'Art. 131, §2o - os Anexos podem ser revistos a cada 120 dias por ato conjunto do Ministro ' +
      'da Fazenda e do Comite Gestor do IBS, ouvido o Ministerio da Saude. Art. 133 - medicamentos ' +
      'registrados na Anvisa. Art. 144 - aplica aos produtos daquela secao o §2o do art. 131. ' +
      'Alteracoes da LC 227/2026 consideradas na versao dos anexos utilizada.');

    b.appendParagraph('3. Metodologia').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    b.appendParagraph(
      'A analise e feita em duas etapas independentes. ETAPA A responde se o produto esta ' +
      'regularizado na Anvisa, por registro, cadastro, notificacao ou isencao. ETAPA B responde ' +
      'se o produto esta enquadrado na hipotese da LC 214, pelo cruzamento NCM x item x anexo. ' +
      'O resultado favoravel exige as duas etapas fechadas. Estar regularizado na Anvisa nao ' +
      'implica, por si so, direito ao tratamento tributario diferenciado.');

    b.appendParagraph('4. Niveis de prova').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    b.appendTable([
      ['Nivel', 'Criterio', 'Uso'],
      ['A - Evidencia direta', 'Registro no XML da NF-e, ou GTIN/codigo exato na base oficial', 'Comprovacao'],
      ['B - Evidencia cruzada', 'Codigo ou descricao forte confirmada por CNPJ do fabricante/fornecedor', 'Comprovacao'],
      ['C - Assistida por IA', 'Candidato real da base escolhido pela IA acima da confianca minima', 'Comprovacao sujeita a revisao por amostragem'],
      ['D - Pendencia', 'Apenas similaridade de nome, ou nenhum candidato', 'NAO utilizada como comprovacao']
    ]).setBorderWidth(0.5);

    b.appendParagraph('5. Resultado').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    b.appendTable([
      ['Indicador', 'Produtos'],
      ['Total analisado', String(d.produtos || 0)],
      ['Com beneficio potencial (enquadrados na LC 214)', String(d.enquadrados || 0)],
      ['Comprovacao suficiente (niveis A e B)', String((d.a || 0) + (d.b || 0))],
      ['Comprovacao assistida por IA (nivel C)', String(d.c || 0)],
      ['Pendentes de conferencia (nivel D)', String(d.dd || 0)],
      ['Enquadrados sem prova de regularizacao (risco)', String(d.risco || 0)]
    ]).setBorderWidth(0.5);

    b.appendParagraph('6. Ressalvas').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    b.appendParagraph('Produtos classificados como PROVAVEL - CONFERIR (nivel D) nao foram ' +
      'considerados como comprovacao definitiva.');
    b.appendParagraph('Correspondencia por similaridade de descricao, isolada, nao fecha ' +
      'veredito em nenhuma hipotese.');
    b.appendParagraph('Etapa A: a base da Anvisa reflete a situacao na data de download informada acima.');
    b.appendParagraph('Etapa B: os Anexos usados sao os da versao informada acima. O art. 131, §2o da ' +
      'LC 214/2025 preve revisao dos Anexos a cada 120 dias por ato conjunto do Ministro da Fazenda e do ' +
      'Comite Gestor do IBS, ouvido o Ministerio da Saude - verificar a publicacao de ato posterior a essa versao.');
    if (d.obs) b.appendParagraph('Observacoes: ' + d.obs);

    doc.saveAndClose();
    return { ok: true, url: doc.getUrl(), nome: nome };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/* ===================== leitura em lote de uma pasta do Drive =====================
 * O navegador nao enxerga caminho de pasta do computador (C:\... ou /Users/...).
 * Duas formas cobrem o caso: seletor de pasta no proprio navegador (Index.html) e,
 * aqui, uma pasta do Google Drive identificada por URL, ID ou nome.
 */
function listarCsvsDrive(ref) {
  try {
    if (!ref) throw new Error('Informe a URL, o ID ou o nome da pasta no Drive.');
    var pasta = null, m = String(ref).match(/[-\w]{25,}/);
    if (m) {
      try { pasta = DriveApp.getFolderById(m[0]); } catch (e) { pasta = null; }
    }
    if (!pasta) {
      var it = DriveApp.getFoldersByName(String(ref).trim());
      if (!it.hasNext()) throw new Error('Pasta nao encontrada: ' + ref);
      pasta = it.next();
    }
    var arquivos = [], f = pasta.getFiles();
    while (f.hasNext()) {
      var a = f.next(), nome = a.getName();
      var tipo = a.getMimeType();
      var ehPlanilha = tipo === MIME_SHEETS || /\.(xls|xlsx)$/i.test(nome);
      if (!/\.(csv|txt)$/i.test(nome) && !ehPlanilha) continue;
      arquivos.push({ id: a.getId(), nome: nome, planilha: ehPlanilha, tamanho: a.getSize(),
                      data: Utilities.formatDate(a.getLastUpdated(), Session.getScriptTimeZone(), 'dd/MM/yyyy') });
      if (arquivos.length >= 300) break;
    }
    return { ok: true, pasta: pasta.getName(), arquivos: arquivos };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/** Le um CSV do Drive. Le em UTF-8 e cai para Windows-1252 se vier com caractere invalido. */
function lerCsvDrive(fileId) {
  try {
    var f = DriveApp.getFileById(fileId);
    var nome = f.getName(), tipo = f.getMimeType();

    if (tipo === MIME_SHEETS) {                       /* ja e uma planilha Google */
      var a = lerSheetsPorId_(fileId, false);
      return { ok: true, matriz: a.linhas, nome: nome, truncado: a.truncado };
    }
    if (/\.(xls|xlsx)$/i.test(nome)) {                /* Excel: converte, le, descarta */
      var id = converterParaSheets_(f.getBlob(), nome);
      var b = lerSheetsPorId_(id, true);
      return { ok: true, matriz: b.linhas, nome: nome, truncado: b.truncado };
    }
    var blob = f.getBlob();
    var txt = blob.getDataAsString('UTF-8');
    if (txt.indexOf('\uFFFD') >= 0) {
      try { txt = blob.getDataAsString('windows-1252'); } catch (e) {}
    }
    return { ok: true, texto: txt, nome: nome };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/* ===================== descoberta automatica da lista da ANVISA =====================
 * O usuario nao precisa saber qual arquivo baixar. Esta funcao varre o diretorio de
 * dados abertos, entra nas pastas promissoras e devolve o melhor candidato para
 * "produtos para saude", com os demais como alternativa.
 */
function pontuarArquivo_(nome) {
  var n = nome.toUpperCase();
  var p = 0;
  if (/PRODUTO/.test(n)) p += 4;
  if (/SAUDE|SA_DE/.test(n)) p += 4;
  if (/CORRELAT/.test(n)) p += 5;
  if (/DISPOSITIV/.test(n)) p += 4;
  if (/REGISTR|CADASTR|NOTIFIC/.test(n)) p += 2;
  if (/MEDICAMENTO/.test(n)) p += 1;
  if (/\.CSV$/.test(n)) p += 2;
  if (/\.ZIP$/.test(n)) p += 1;
  if (/CONSULTA_PUBLICA|FISCALIZA|INSPECAO|PRECO|CMED|ALIMENTO|COSMETIC|SANEANTE|AGROTOX/.test(n)) p -= 6;
  return p;
}

function descobrirBaseAnvisa() {
  try {
    var raiz = 'https://dados.anvisa.gov.br/dados/';
    var vistos = {}, arquivos = [], pastas = [raiz], nivel = 0;

    while (pastas.length && nivel < 2) {
      var proxima = [];
      for (var i = 0; i < pastas.length && i < 8; i++) {
        var r = listarArquivosAnvisa(pastas[i]);
        if (!r.ok) continue;
        r.arquivos.forEach(function (a) {
          if (vistos[a]) return;
          vistos[a] = 1;
          var nome = a.replace(/\/$/, '').split('/').pop();
          if (/\/$/.test(a)) {
            if (pontuarArquivo_(nome) > 2) proxima.push(a);
          } else {
            arquivos.push({ url: a, nome: nome, ponto: pontuarArquivo_(nome) });
          }
        });
      }
      pastas = proxima; nivel++;
    }

    arquivos.sort(function (a, b) { return b.ponto - a.ponto; });
    var bons = arquivos.filter(function (a) { return a.ponto >= 6; });
    return {
      ok: true,
      escolhido: bons.length ? bons[0] : null,
      candidatos: arquivos.slice(0, 25)
    };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/* ===================== conversao de planilha legada =====================
 * O Dominio exporta um .xls em BIFF antigo e mal formado: o leitor do navegador
 * devolve zero linhas e ate bibliotecas de desktop recusam o arquivo. O Drive do
 * Google recupera, entao a conversao roda aqui, no servidor.
 */
var MIME_SHEETS = 'application/vnd.google-apps.spreadsheet';

function converterParaSheets_(blob, nome) {
  var limite = '----planilha' + Date.now();
  var meta = { name: nome || 'conversao', mimeType: MIME_SHEETS };
  var topo = Utilities.newBlob(
    '--' + limite + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(meta) + '\r\n--' + limite + '\r\n' +
    'Content-Type: ' + (blob.getContentType() || 'application/vnd.ms-excel') + '\r\n\r\n'
  ).getBytes();
  var fim = Utilities.newBlob('\r\n--' + limite + '--\r\n').getBytes();
  var corpo = topo.concat(blob.getBytes()).concat(fim);

  var r = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true', {
      method: 'post',
      contentType: 'multipart/related; boundary=' + limite,
      payload: corpo,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
  if (r.getResponseCode() !== 200) {
    throw new Error('Drive recusou a conversao: HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
  }
  return JSON.parse(r.getContentText()).id;
}

/**
 * Le a primeira aba de uma planilha do Drive.
 *
 * Duas precaucoes que fazem diferenca em arquivo grande:
 * 1. le em blocos, para nao estourar o tempo de execucao;
 * 2. descarta as colunas totalmente vazias antes de devolver. O relatorio do
 *    Dominio tem 80 colunas por causa de celulas mescladas, das quais ~20 tem
 *    dado. Devolver as 80 significaria mandar mais de um milhao de celulas para
 *    o navegador - e e isso que trava a tela.
 */
function lerSheetsPorId_(id, limparDepois) {
  var MAX_LIN = 40000, MAX_COL = 120, BLOCO = 2000;
  try {
    var aba = SpreadsheetApp.openById(id).getSheets()[0];
    var nl = Math.min(aba.getLastRow(), MAX_LIN);
    var nc = Math.min(aba.getLastColumn(), MAX_COL);
    if (nl < 1 || nc < 1) return { linhas: [], truncado: false };

    var linhas = [], usada = [];
    for (var i = 0; i < nc; i++) usada.push(false);

    for (var ini = 1; ini <= nl; ini += BLOCO) {
      var qtd = Math.min(BLOCO, nl - ini + 1);
      var bloco = aba.getRange(ini, 1, qtd, nc).getDisplayValues();
      for (var r = 0; r < bloco.length; r++) {
        var l = bloco[r], vazia = true;
        for (var c = 0; c < nc; c++) {
          if (l[c] !== '' && l[c] != null) { usada[c] = true; vazia = false; }
        }
        if (!vazia) linhas.push(l);
      }
    }

    var cols = [];
    for (var k = 0; k < nc; k++) if (usada[k]) cols.push(k);
    var saida = [];
    for (var j = 0; j < linhas.length; j++) {
      var nova = [];
      for (var m = 0; m < cols.length; m++) nova.push(linhas[j][cols[m]]);
      saida.push(nova);
    }
    return { linhas: saida, truncado: aba.getLastRow() > MAX_LIN, colunas: cols.length, colunasOriginais: nc };
  } finally {
    if (limparDepois) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
  }
}

/**
 * Recebe o arquivo do navegador em base64, converte no Drive, le e apaga a copia.
 * Usado so quando a leitura no navegador falha.
 */
var LIMITE_UPLOAD = 6 * 1024 * 1024;   /* acima disso o navegador trava ao serializar */

function converterPlanilhaEnviada(base64, nome, mime) {
  try {
    if (base64 && base64.length > LIMITE_UPLOAD * 1.4) {
      throw new Error('Arquivo grande demais para subir pelo navegador. ' +
        'Coloque o arquivo numa pasta do Google Drive e use o campo de pasta no passo 1.');
    }
    var bytes = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(bytes, mime || 'application/vnd.ms-excel', nome || 'arquivo.xls');
    var id = converterParaSheets_(blob, nome);
    var m = lerSheetsPorId_(id, true);
    if (!m.linhas.length) throw new Error('A planilha convertida veio vazia.');
    return { ok: true, linhas: m.linhas, nome: nome, truncado: m.truncado };
  } catch (e) {
    return { ok: false, erro: String(e) };
  }
}

/* ===================== CARGA AUTOMATICA DA LISTA DA ANVISA =====================
 * O usuario nao precisa clicar em nada. Na primeira vez que a ferramenta abre, a
 * lista oficial e baixada em segundo plano, fica guardada na aba BASE_ANVISA e
 * se atualiza sozinha todo mes.
 *
 * O arquivo tem ~27 MB e nao cabe numa unica execucao do Apps Script, entao a
 * carga e feita em pedacos, por um gatilho que roda de minuto em minuto ate
 * terminar. Cada pedaco e cortado no ultimo fim de linha em BYTES, para nao
 * quebrar um caractere acentuado no meio.
 */
var ANVISA_CSV = 'https://dados.anvisa.gov.br/dados/TA_PRODUTO_SAUDE_SITE.csv';
var PEDACO_BYTES = 4 * 1024 * 1024;
var CAB_ANVISA = ['REGISTRO', 'PRODUTO', 'EMPRESA', 'CNPJ', 'SITUACAO', 'VENCIMENTO'];

/* apelidos de coluna: a Anvisa ja mudou esses nomes mais de uma vez */
var ALIAS_ANVISA = {
  REGISTRO:   ['NU_REGISTRO', 'REGISTRO', 'NUMERO_REGISTRO', 'NU_PROCESSO', 'NUMERO_DO_REGISTRO'],
  PRODUTO:    ['NO_PRODUTO', 'NOME_COMERCIAL', 'NO_NOME_COMERCIAL', 'NOME_TECNICO', 'NO_NOME_TECNICO', 'PRODUTO', 'NOME'],
  EMPRESA:    ['NO_RAZAO_SOCIAL', 'RAZAO_SOCIAL', 'EMPRESA', 'DETENTOR', 'FABRICANTE', 'NO_EMPRESA'],
  CNPJ:       ['NU_CNPJ', 'CNPJ', 'NU_CNPJ_EMPRESA'],
  SITUACAO:   ['DS_SITUACAO', 'SITUACAO', 'ST_SITUACAO', 'SITUACAO_REGISTRO'],
  VENCIMENTO: ['DT_VENCIMENTO', 'VENCIMENTO', 'DT_VALIDADE', 'VALIDADE']
};

function props_() { return PropertiesService.getDocumentProperties(); }
function estadoCarga_() { try { return JSON.parse(props_().getProperty('cargaAnvisa') || 'null'); } catch (e) { return null; } }
function gravarEstado_(e) { props_().setProperty('cargaAnvisa', JSON.stringify(e)); }

function semAcentoMaiusc_(s) {
  return String(s || '').toUpperCase()
    .replace(/[ÁÀÂÃÄ]/g, 'A').replace(/[ÉÈÊË]/g, 'E').replace(/[ÍÌÎÏ]/g, 'I')
    .replace(/[ÓÒÔÕÖ]/g, 'O').replace(/[ÚÙÛÜ]/g, 'U').replace(/Ç/g, 'C')
    .replace(/[^A-Z0-9]/g, '_');
}

/** Descobre, no cabecalho do CSV da Anvisa, onde esta cada coluna que interessa. */
function mapearColunas_(cabecalho) {
  var norm = cabecalho.map(semAcentoMaiusc_), mapa = {};
  for (var campo in ALIAS_ANVISA) {
    mapa[campo] = -1;
    var lista = ALIAS_ANVISA[campo];
    for (var a = 0; a < lista.length && mapa[campo] < 0; a++) {
      for (var i = 0; i < norm.length; i++) {
        if (norm[i] === lista[a]) { mapa[campo] = i; break; }
      }
    }
    if (mapa[campo] < 0) {                       /* segunda passada: por conteudo */
      for (var b = 0; b < lista.length && mapa[campo] < 0; b++) {
        for (var j = 0; j < norm.length; j++) {
          if (norm[j].indexOf(lista[b]) >= 0) { mapa[campo] = j; break; }
        }
      }
    }
  }
  return mapa;
}

function partirLinha_(linha, sep) {
  var out = [], campo = '', aspas = false;
  for (var i = 0; i < linha.length; i++) {
    var c = linha.charAt(i);
    if (aspas) {
      if (c === '"') { if (linha.charAt(i + 1) === '"') { campo += '"'; i++; } else aspas = false; }
      else campo += c;
    } else if (c === '"') aspas = true;
    else if (c === sep) { out.push(campo); campo = ''; }
    else campo += c;
  }
  out.push(campo);
  return out;
}

/** Estado da carga, para a tela mostrar o progresso sem o usuario perguntar. */
function statusCargaAnvisa() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss ? ss.getSheetByName(ABA_BASE) : null;
    var linhas = aba ? Math.max(aba.getLastRow() - 1, 0) : 0;
    var e = estadoCarga_();
    if (e && !e.fim) {
      return { ok: true, estado: 'carregando', linhas: e.linhas || 0,
               pct: e.total ? Math.round(100 * e.offset / e.total) : 0,
               mensagem: 'Baixando a lista oficial da Anvisa em segundo plano' };
    }
    if (linhas > 0) {
      return { ok: true, estado: 'pronta', linhas: linhas,
               data: props_().getProperty('dataBaseAnvisa') || '',
               mensagem: linhas + ' produtos na lista oficial' };
    }
    return { ok: true, estado: 'vazia', linhas: 0, erro: (e && e.erro) || '' };
  } catch (e2) { return { ok: false, erro: String(e2) }; }
}

/** Chamada pela tela ao abrir: se nao houver lista, comeca a baixar sozinha. */
function garantirBaseAnvisa() {
  var st = statusCargaAnvisa();
  if (st.estado === 'vazia') {
    try { iniciarCargaAnvisa(); return statusCargaAnvisa(); }
    catch (e) { return { ok: false, estado: 'erro', erro: String(e) }; }
  }
  return st;
}

function iniciarCargaAnvisa() {
  limparGatilho_('passoCargaAnvisa');
  var aba = aba_(ABA_BASE, CAB_ANVISA);
  aba.clear();
  aba.getRange(1, 1, 1, CAB_ANVISA.length).setValues([CAB_ANVISA])
     .setFontWeight('bold').setBackground('#14432F').setFontColor('#FFFFFF');
  aba.setFrozenRows(1);
  gravarEstado_({ offset: 0, linhas: 0, total: 0, mapa: null, sep: ';', fim: false, erro: '' });
  ScriptApp.newTrigger('passoCargaAnvisa').timeBased().everyMinutes(1).create();
  passoCargaAnvisa();
  return { ok: true };
}

function limparGatilho_(nome) {
  var gs = ScriptApp.getProjectTriggers();
  for (var i = 0; i < gs.length; i++) {
    if (gs[i].getHandlerFunction() === nome) ScriptApp.deleteTrigger(gs[i]);
  }
}

/** Um pedaco por execucao. O gatilho chama de novo ate acabar. */
function passoCargaAnvisa() {
  var e = estadoCarga_();
  if (!e || e.fim) { limparGatilho_('passoCargaAnvisa'); return; }
  var trava = LockService.getDocumentLock();
  if (!trava.tryLock(1000)) return;                 /* outro passo ja esta rodando */
  try {
    var ini = e.offset, fim = ini + PEDACO_BYTES - 1;
    var r = UrlFetchApp.fetch(ANVISA_CSV, {
      headers: { Range: 'bytes=' + ini + '-' + fim },
      muteHttpExceptions: true, followRedirects: true
    });
    var cod = r.getResponseCode();
    if (cod !== 200 && cod !== 206) throw new Error('HTTP ' + cod + ' ao baixar a lista da Anvisa');

    var faixa = r.getHeaders()['Content-Range'] || r.getHeaders()['content-range'] || '';
    if (faixa) {
      var t = faixa.split('/')[1];
      if (t && !isNaN(t)) e.total = parseInt(t, 10);
    }
    var bytes = r.getBlob().getBytes();
    var parcial = (cod === 206) && (!e.total || ini + bytes.length < e.total);

    /* corta no ultimo fim de linha, em BYTES, para nao partir acentuacao */
    var corte = bytes.length;
    if (parcial) {
      for (var k = bytes.length - 1; k >= 0; k--) { if (bytes[k] === 10) { corte = k + 1; break; } }
      if (corte === bytes.length) throw new Error('Pedaco sem fim de linha - arquivo em formato inesperado.');
    }
    var texto = Utilities.newBlob(bytes.slice(0, corte)).getDataAsString('UTF-8');
    if (texto.indexOf('�') >= 0) {
      texto = Utilities.newBlob(bytes.slice(0, corte)).getDataAsString('ISO-8859-1');
    }

    var linhas = texto.split(/\r?\n/);
    if (ini === 0) {
      var cabTxt = linhas.shift() || '';
      e.sep = (cabTxt.split(';').length > cabTxt.split(',').length) ? ';' : ',';
      e.mapa = mapearColunas_(partirLinha_(cabTxt, e.sep));
      if (e.mapa.PRODUTO < 0) throw new Error('Nao achei a coluna de nome do produto no arquivo da Anvisa.');
    }

    var novas = [];
    for (var i = 0; i < linhas.length; i++) {
      if (!linhas[i]) continue;
      var c = partirLinha_(linhas[i], e.sep);
      var nome = e.mapa.PRODUTO >= 0 ? (c[e.mapa.PRODUTO] || '').trim() : '';
      if (!nome) continue;
      novas.push([
        e.mapa.REGISTRO   >= 0 ? (c[e.mapa.REGISTRO]   || '').trim() : '',
        nome,
        e.mapa.EMPRESA    >= 0 ? (c[e.mapa.EMPRESA]    || '').trim() : '',
        e.mapa.CNPJ       >= 0 ? (c[e.mapa.CNPJ]       || '').trim() : '',
        e.mapa.SITUACAO   >= 0 ? (c[e.mapa.SITUACAO]   || '').trim() : '',
        e.mapa.VENCIMENTO >= 0 ? (c[e.mapa.VENCIMENTO] || '').trim() : ''
      ]);
    }
    if (novas.length) {
      var aba = aba_(ABA_BASE, CAB_ANVISA);
      aba.getRange(aba.getLastRow() + 1, 1, novas.length, CAB_ANVISA.length).setValues(novas);
      e.linhas = (e.linhas || 0) + novas.length;
    }

    e.offset = ini + corte;
    if (!parcial || (e.total && e.offset >= e.total)) {
      e.fim = true;
      limparGatilho_('passoCargaAnvisa');
      props_().setProperty('dataBaseAnvisa', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'));
      registrarCargaBase({ linhas: e.linhas, arquivo: 'TA_PRODUTO_SAUDE_SITE.csv', hash: '' });
      agendarAtualizacaoMensal();
    }
    gravarEstado_(e);
  } catch (err) {
    var e2 = estadoCarga_() || {};
    e2.erro = String(err); e2.fim = true;
    gravarEstado_(e2);
    limparGatilho_('passoCargaAnvisa');
  } finally {
    trava.releaseLock();
  }
}

/** Uma vez por mes a lista se atualiza sozinha, de madrugada. */
function agendarAtualizacaoMensal() {
  limparGatilho_('atualizarBaseAnvisaMensal');
  ScriptApp.newTrigger('atualizarBaseAnvisaMensal').timeBased()
    .onMonthDay(5).atHour(3).create();
}

function atualizarBaseAnvisaMensal() { iniciarCargaAnvisa(); }