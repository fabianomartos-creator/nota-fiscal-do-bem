/**
 * Nota Fiscal do Bem — recebe os cadastros da landing page e grava na planilha.
 *
 * Onde usar: Planilha "Inscrições - Nota Fiscal do Bem" > Extensões > Apps Script
 * (arquivo Código.gs). Depois de colar, é preciso publicar uma NOVA VERSÃO da
 * implantação (Implantar > Gerenciar implantações > editar > Nova versão).
 *
 * Este arquivo no GitHub é só a cópia versionada: ele NÃO é implantado sozinho.
 */

// Ordem das colunas numa planilha nova. Numa planilha que já tem dados, as
// colunas que faltarem são acrescentadas à direita, sem mover as existentes.
const CABECALHO = [
  'Data/Hora', 'Nome', 'CPF', 'WhatsApp', 'E-mail',
  'utm_source', 'utm_medium', 'utm_campaign', 'Página',
  'Observação', 'Captador',
  'CEP', 'Endereço', 'Número', 'Bairro', 'Cidade', 'UF'
];

// Colunas que mudaram de nome: o cabeçalho antigo é renomeado NO LUGAR (os dados da coluna ficam intactos).
const RENOMEAR = { 'Cadastrador': 'Captador' };

// Colunas gravadas como texto (para não perder zeros nem virar número/data).
const COLUNAS_TEXTO = ['CPF', 'WhatsApp', 'CEP', 'Número'];

// Fuso horário de Brasília (sem horário de verão desde 2019).
const FUSO = 'America/Sao_Paulo';
const FORMATO_DATA = 'dd/MM/yyyy HH:mm:ss';

// Tamanho máximo aceito por campo (o formulário já limita; aqui é a garantia).
const LIMITE_PADRAO = 200;
const LIMITE_OBSERVACAO = 300;

// Abre a página quando alguém acessa o link /exec (variante opcional, requer
// um arquivo HTML chamado "Pagina" no projeto do Apps Script).
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Pagina');
  t.params = JSON.stringify((e && e.parameter) || {}).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('Nota Fiscal do Bem | Seja Notável')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Chamado pela variante servida pelo próprio Google (google.script.run).
function salvar(d) {
  gravar_(d || {});
  return true;
}

// Chamado pela landing page do GitHub Pages (index.html).
function doPost(e) {
  gravar_((e && e.parameter) || {});
  return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
}

// Evita que o texto vire fórmula na planilha (=, +, -, @, TAB, CR no início,
// mesmo depois de espaços ou quebras de linha). O conteúdo é mantido como veio.
function limpo_(v, max) {
  v = String(v == null ? '' : v).slice(0, max || LIMITE_PADRAO);
  const semEspacoInicial = v.replace(/^[ \n]+/, '');
  return /^[=+\-@\t\r]/.test(semEspacoInicial) ? "'" + v : v;
}

// Garante que a linha 1 tenha todas as colunas de CABECALHO.
// - Planilha vazia (ou linha 1 vazia): escreve o cabeçalho completo.
// - Planilha com cabeçalho: só acrescenta à direita as colunas que faltam.
// Nunca apaga, sobrescreve ou reordena colunas/dados existentes.
// Retorna a lista de títulos da linha 1, na ordem em que estão na planilha.
function garantirCabecalho_(sh) {
  const ultimaColuna = sh.getLastColumn();
  let atual = [];
  if (sh.getLastRow() > 0 && ultimaColuna > 0) {
    atual = sh.getRange(1, 1, 1, ultimaColuna).getValues()[0].map(function (v) {
      return String(v).trim();
    });
  }
  const linha1Vazia = atual.every(function (v) { return v === ''; });

  let criouAgora = false;
  if (linha1Vazia) {
    sh.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]);
    atual = CABECALHO.slice();
    criouAgora = true;
  } else {
    // renomeia cabeçalhos antigos (só o texto da célula da linha 1; nada mais é tocado)
    atual.forEach(function (h, i) {
      const novo = RENOMEAR[h];
      if (novo && atual.indexOf(novo) === -1) {
        sh.getRange(1, i + 1).setValue(novo);
        atual[i] = novo;
      }
    });
    // remove células vazias do fim da linha 1 para acrescentar logo depois do último título
    while (atual.length && atual[atual.length - 1] === '') atual.pop();
    const faltando = CABECALHO.filter(function (h) { return atual.indexOf(h) === -1; });
    if (faltando.length) {
      sh.getRange(1, atual.length + 1, 1, faltando.length).setValues([faltando]);
      // colunas novas que devem ser texto (ex.: CEP, Número) já nascem formatadas como texto
      faltando.forEach(function (h, i) {
        if (COLUNAS_TEXTO.indexOf(h) !== -1) {
          sh.getRange(1, atual.length + 1 + i, sh.getMaxRows(), 1).setNumberFormat('@');
        }
      });
      atual = atual.concat(faltando);
    }
  }

  if (criouAgora || sh.getFrozenRows() === 0) {
    sh.setFrozenRows(1);
    COLUNAS_TEXTO.forEach(function (nome) {
      const i = atual.indexOf(nome);
      if (i !== -1) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  }
  return atual;
}

function gravar_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getSpreadsheetTimeZone() !== FUSO) ss.setSpreadsheetTimeZone(FUSO);
    const sh = ss.getSheets()[0]; // primeira aba
    const cabecalho = garantirCabecalho_(sh);

    const valores = {
      // Data/hora já escrita no horário de Brasília (não depende do fuso da planilha nem da conta Google)
      'Data/Hora': Utilities.formatDate(new Date(), FUSO, 'yyyy-MM-dd HH:mm:ss'),
      'Nome': limpo_(p.nome),
      'CPF': limpo_(p.cpf),
      'WhatsApp': limpo_(p.whatsapp),
      'E-mail': limpo_(p.email),
      'utm_source': limpo_(p.utm_source),
      'utm_medium': limpo_(p.utm_medium),
      'utm_campaign': limpo_(p.utm_campaign),
      'Página': limpo_(p.pagina),
      'Observação': limpo_(p.observacao, LIMITE_OBSERVACAO),
      // aceita "cadastrador" para páginas antigas ainda em cache no navegador
      'Captador': limpo_(p.captador || p.cadastrador),
      'CEP': limpo_(p.cep),
      'Endereço': limpo_(p.endereco),
      'Número': limpo_(p.numero),
      'Bairro': limpo_(p.bairro),
      'Cidade': limpo_(p.cidade),
      'UF': limpo_(p.uf)
    };

    // Monta a linha na ordem real das colunas da planilha.
    const linha = cabecalho.map(function (h) {
      return Object.prototype.hasOwnProperty.call(valores, h) ? valores[h] : '';
    });
    sh.appendRow(linha);
    const iData = cabecalho.indexOf('Data/Hora');
    if (iData !== -1) sh.getRange(sh.getLastRow(), iData + 1).setNumberFormat(FORMATO_DATA);
  } finally {
    lock.releaseLock();
  }
}

// Atualiza só o cabeçalho (renomeia Cadastrador -> Captador, acrescenta as colunas novas)
// e coloca a planilha no fuso de Brasília. Serve para conferir se este código está no
// projeto certo: selecione esta função na barra de cima do editor e clique em Executar.
// Não grava nenhuma inscrição e não altera os dados existentes.
function atualizarCabecalho() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSpreadsheetTimeZone() !== FUSO) ss.setSpreadsheetTimeZone(FUSO);
  garantirCabecalho_(ss.getSheets()[0]);
}

// =====================================================================
// CORREÇÃO ÚNICA DOS HORÁRIOS ANTIGOS
// As inscrições gravadas ANTES desta versão ficaram 5 horas atrasadas (a planilha
// estava no fuso da Califórnia). Estas duas funções corrigem só a coluna Data/Hora
// dessas linhas. Use na ordem: 1) previaCorrecaoHorarios  2) corrigirHorariosAntigos
// Rode logo depois de publicar a Nova versão. Só funciona UMA vez (depois ela se bloqueia).
// =====================================================================
const DESLOCAMENTO_HORAS = 5;
// Linhas com data mais recente que (agora - 2,5 h) já foram gravadas no horário certo
// pela versão nova e NÃO são alteradas. Por isso, rode a correção em até 2 h após publicar.
const JANELA_SEGURANCA_HORAS = 2.5;
const MARCA_CORRECAO = 'HORARIOS_CORRIGIDOS';

// 'dd/MM/yyyy HH:mm:ss' (como aparece na célula) -> milissegundos do relógio de parede (sem fuso)
function brParaMs_(t) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})$/.exec(String(t).trim());
  return m ? Date.UTC(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], +m[6]) : null;
}
// 'yyyy-MM-dd HH:mm:ss' -> milissegundos do relógio de parede
function isoParaMs_(t) {
  const m = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/.exec(t);
  return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
}
function pad2_(n) { return ('0' + n).slice(-2); }
function msParaIso_(ms) {
  const d = new Date(ms);
  return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate()) + ' ' +
         pad2_(d.getUTCHours()) + ':' + pad2_(d.getUTCMinutes()) + ':' + pad2_(d.getUTCSeconds());
}
function msParaBr_(ms) {
  const d = new Date(ms);
  return pad2_(d.getUTCDate()) + '/' + pad2_(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear() + ' ' +
         pad2_(d.getUTCHours()) + ':' + pad2_(d.getUTCMinutes()) + ':' + pad2_(d.getUTCSeconds());
}

// Monta o plano sem escrever nada.
function planejarCorrecao_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const ultima = sh.getLastRow();
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (v) { return String(v).trim(); });
  const col = cab.indexOf('Data/Hora') + 1;
  if (!col) throw new Error('Coluna "Data/Hora" não encontrada na linha 1. Nada foi alterado.');
  const plano = { sh: sh, col: col, itens: [], ignoradas: 0, vazias: 0, problemas: [] };
  if (ultima < 2) return plano;

  const exibidas = sh.getRange(2, col, ultima - 1, 1).getDisplayValues();
  const agoraMs = isoParaMs_(Utilities.formatDate(new Date(), FUSO, 'yyyy-MM-dd HH:mm:ss'));
  const limiteMs = agoraMs - JANELA_SEGURANCA_HORAS * 3600000;

  exibidas.forEach(function (linha, i) {
    const texto = String(linha[0]).trim();
    const numLinha = i + 2;
    if (texto === '') { plano.vazias++; return; }
    const ms = brParaMs_(texto);
    if (ms === null) { plano.problemas.push('linha ' + numLinha + ': "' + texto + '"'); return; }
    if (ms >= limiteMs) { plano.ignoradas++; return; }
    plano.itens.push({ linha: numLinha, antes: texto, depois: ms + DESLOCAMENTO_HORAS * 3600000 });
  });
  return plano;
}

// PASSO 1: só mostra o que seria alterado. Não escreve nada.
// Veja o resultado em "Registro de execução" (embaixo no editor).
function previaCorrecaoHorarios() {
  const plano = planejarCorrecao_();
  Logger.log('PRÉVIA — nada foi alterado.');
  Logger.log('Linhas que seriam corrigidas (+' + DESLOCAMENTO_HORAS + ' h): ' + plano.itens.length);
  Logger.log('Linhas ignoradas por já estarem no horário certo: ' + plano.ignoradas);
  if (plano.problemas.length) {
    Logger.log('ATENÇÃO: datas que não consegui ler (a correção NÃO vai rodar enquanto existirem): ' + plano.problemas.join(' | '));
  }
  const mostrar = plano.itens.length <= 6 ? plano.itens : plano.itens.slice(0, 3).concat(plano.itens.slice(-3));
  mostrar.forEach(function (it) {
    Logger.log('linha ' + it.linha + ': ' + it.antes + '  ->  ' + msParaBr_(it.depois));
  });
  if (plano.itens.length > 6) Logger.log('(mostrando as 3 primeiras e as 3 últimas)');
  const marca = PropertiesService.getScriptProperties().getProperty(MARCA_CORRECAO);
  if (marca) Logger.log('AVISO: a correção já foi feita em ' + marca + '. Rodar de novo será bloqueado.');
}

// PASSO 2: aplica a correção, uma única vez. Só altera a coluna Data/Hora.
function corrigirHorariosAntigos() {
  const props = PropertiesService.getScriptProperties();
  const jaFeito = props.getProperty(MARCA_CORRECAO);
  if (jaFeito) throw new Error('Os horários antigos já foram corrigidos em ' + jaFeito + '. Nada foi alterado.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const plano = planejarCorrecao_();
    if (plano.problemas.length) {
      throw new Error('Há datas que não consegui ler, então NADA foi alterado: ' + plano.problemas.join(' | '));
    }
    if (!plano.itens.length) throw new Error('Nenhuma linha para corrigir. Nada foi alterado.');

    // grava em blocos de linhas seguidas
    let i = 0;
    while (i < plano.itens.length) {
      let j = i;
      while (j + 1 < plano.itens.length && plano.itens[j + 1].linha === plano.itens[j].linha + 1) j++;
      const bloco = plano.itens.slice(i, j + 1);
      const faixa = plano.sh.getRange(bloco[0].linha, plano.col, bloco.length, 1);
      faixa.setValues(bloco.map(function (it) { return [msParaIso_(it.depois)]; }));
      faixa.setNumberFormat(FORMATO_DATA);
      i = j + 1;
    }
    props.setProperty(MARCA_CORRECAO, Utilities.formatDate(new Date(), FUSO, 'yyyy-MM-dd HH:mm:ss'));
    Logger.log('Pronto: ' + plano.itens.length + ' horários corrigidos (+' + DESLOCAMENTO_HORAS + ' h). Ignoradas (já certas): ' + plano.ignoradas + '.');
  } finally {
    lock.releaseLock();
  }
}
