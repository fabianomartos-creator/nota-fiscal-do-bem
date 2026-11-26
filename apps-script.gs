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
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // primeira aba
    const cabecalho = garantirCabecalho_(sh);

    const valores = {
      'Data/Hora': new Date(),
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
  } finally {
    lock.releaseLock();
  }
}

// Atualiza só o cabeçalho (renomeia Cadastrador -> Captador e acrescenta as colunas novas).
// Serve para conferir se este código está no projeto certo: selecione esta função
// na barra de cima do editor e clique em Executar. Não grava nenhuma inscrição.
function atualizarCabecalho() {
  garantirCabecalho_(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]);
}
