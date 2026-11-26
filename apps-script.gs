/**
 * Nota Fiscal do Bem — recebe os cadastros dos captadores e grava na planilha
 * "Cadastros - Nota Fiscal do Bem".
 *
 * Onde usar: nessa planilha > Extensões > Apps Script (arquivo Código.gs).
 * Ordem: cole este código > Salvar > Executar "prepararLogin" (autoriza e prepara a planilha)
 *        > Implantar > Nova implantação > App da Web (Executar como: Eu; Acesso: Qualquer pessoa).
 *
 * Este arquivo no GitHub é só a cópia versionada: ele NÃO é implantado sozinho.
 * (O código da planilha antiga está em legado/apps-script-planilha-antiga.gs.)
 *
 * Regras:
 * - Só cadastra quem entrou com login e senha (cada captador tem os seus).
 * - As senhas ficam na aba "Captadores" SÓ como hash (nunca em texto).
 * - O nome gravado na coluna Captador vem sempre da aba "Captadores", nunca da página.
 * - Depois de MAX_FALHAS senhas erradas, o login fica bloqueado por BLOQUEIO_SEGUNDOS.
 */

const CABECALHO = ['Data/Hora', 'Nome', 'CPF', 'WhatsApp', 'Captador'];

// Colunas gravadas como texto (para não perder zeros nem virar número/data).
const COLUNAS_TEXTO = ['CPF', 'WhatsApp'];

// Fuso horário de Brasília (sem horário de verão desde 2019).
const FUSO = 'America/Sao_Paulo';
const FORMATO_DATA = 'dd/MM/yyyy HH:mm:ss';

// Tamanho máximo aceito por campo (a página já limita; aqui é a garantia).
const LIMITE_PADRAO = 200;

const ABA_CAPTADORES = 'Captadores';
const CAB_CAPTADORES = ['Login', 'Nome', 'Hash da senha', 'Salt', 'Ativo'];
const MAX_FALHAS = 5;
const BLOQUEIO_SEGUNDOS = 900;
const SESSAO_HORAS = 12;
const PROP_SEGREDO = 'SEGREDO_SESSAO';

// =====================================================================
// ENTRADA (o que a página chama)
// =====================================================================

// Abrir o link /exec no navegador mostra este aviso (serve só para conferir que o serviço está no ar).
function doGet() {
  return json_({ ok: true, servico: 'Nota Fiscal do Bem' });
}

// Recebe tudo que a página manda. Sempre responde em JSON (a página lê a resposta).
function doPost(e) {
  let resposta;
  try {
    resposta = rotear_((e && e.parameter) || {});
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    resposta = { ok: false, erro: 'servidor' };
  }
  return json_(resposta);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function rotear_(p) {
  const acao = String(p.acao || '');
  if (acao === 'login') return acaoLogin_(p);
  if (acao === 'cadastrar') return acaoCadastrar_(p);
  return { ok: false, erro: 'acao_invalida' };
}

// =====================================================================
// LOGIN
// =====================================================================

function bytesParaHex_(bytes) {
  return bytes.map(function (b) { return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2); }).join('');
}
function hashSenha_(salt, senha) {
  return bytesParaHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + senha, Utilities.Charset.UTF_8));
}
// Compara sem parar no primeiro caractere diferente.
function iguais_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}
function segredo_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty(PROP_SEGREDO);
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty(PROP_SEGREDO, s); }
  return s;
}
function assinar_(texto) {
  return bytesParaHex_(Utilities.computeHmacSha256Signature(texto, segredo_()));
}
function criarToken_(login) {
  const validoAte = Date.now() + SESSAO_HORAS * 3600000;
  const corpo = Utilities.base64EncodeWebSafe(login + '|' + validoAte, Utilities.Charset.UTF_8);
  return corpo + '.' + assinar_(corpo);
}
// Retorna o captador (ativo) dono do token, ou null.
function validarToken_(token) {
  const partes = String(token || '').split('.');
  if (partes.length !== 2 || !partes[0] || !partes[1]) return null;
  if (!iguais_(assinar_(partes[0]), partes[1])) return null;
  let texto;
  try {
    texto = Utilities.newBlob(Utilities.base64DecodeWebSafe(partes[0])).getDataAsString('UTF-8');
  } catch (e) { return null; }
  const i = texto.lastIndexOf('|');
  const login = texto.slice(0, i);
  const validoAte = Number(texto.slice(i + 1));
  if (!login || !(validoAte > Date.now())) return null;
  const u = buscarCaptador_(login);
  return u && u.ativo ? u : null;
}

function abaCaptadores_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_CAPTADORES);
  if (!sh && criar) {
    sh = ss.insertSheet(ABA_CAPTADORES, ss.getSheets().length); // sempre no fim: a 1ª aba continua sendo a dos cadastros
    sh.getRange('A:E').setNumberFormat('@');
    sh.getRange(1, 1, 1, CAB_CAPTADORES.length).setValues([CAB_CAPTADORES]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function buscarCaptador_(login) {
  const sh = abaCaptadores_(false);
  if (!sh || sh.getLastRow() < 2) return null;
  const linhas = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (let i = 0; i < linhas.length; i++) {
    if (String(linhas[i][0]).trim().toLowerCase() === login) {
      return {
        linha: i + 2, login: login, nome: String(linhas[i][1]).trim(),
        hash: String(linhas[i][2]), salt: String(linhas[i][3]),
        ativo: String(linhas[i][4]).trim().toUpperCase() === 'SIM'
      };
    }
  }
  return null;
}

function acaoLogin_(p) {
  const login = String(p.login || '').trim().toLowerCase().slice(0, 60);
  const senha = String(p.senha || '').slice(0, 200);
  if (!login || !senha) return { ok: false, erro: 'dados' };

  const cache = CacheService.getScriptCache();
  const chave = 'falhas_' + login;
  const falhas = Number(cache.get(chave) || 0);
  if (falhas >= MAX_FALHAS) return { ok: false, erro: 'bloqueado' };

  const u = buscarCaptador_(login);
  const hashCalculado = hashSenha_(u ? u.salt : 'sem-usuario', senha); // calcula sempre, existindo ou não
  if (!u || !u.ativo || !iguais_(hashCalculado, u.hash)) {
    cache.put(chave, String(falhas + 1), BLOQUEIO_SEGUNDOS);
    return { ok: false, erro: 'credenciais' };
  }
  cache.remove(chave);
  return { ok: true, token: criarToken_(login), nome: u.nome };
}

// =====================================================================
// CADASTRO
// =====================================================================

// Evita que o texto vire fórmula na planilha (=, +, -, @, TAB, CR no início,
// mesmo depois de espaços ou quebras de linha). O conteúdo é mantido como veio.
function limpo_(v, max) {
  v = String(v == null ? '' : v).slice(0, max || LIMITE_PADRAO);
  const semEspacoInicial = v.replace(/^[ \n]+/, '');
  return /^[=+\-@\t\r]/.test(semEspacoInicial) ? "'" + v : v;
}

function acaoCadastrar_(p) {
  const u = validarToken_(p.token);
  if (!u) return { ok: false, erro: 'sessao' };
  const nome = String(p.nome || '').trim();
  const whatsapp = String(p.whatsapp || '').trim();
  if (nome.split(/\s+/).filter(Boolean).length < 2) return { ok: false, erro: 'nome' };
  if (whatsapp.replace(/\D/g, '').length < 10) return { ok: false, erro: 'whatsapp' };
  const idem = String(p.idem || '');
  const chaveIdem = /^[A-Za-z0-9-]{8,64}$/.test(idem) ? 'idem_' + idem : null;
  const gravou = gravar_({
    nome: nome, cpf: p.cpf, whatsapp: whatsapp,
    captador: u.nome            // vem da aba Captadores, não da página
  }, chaveIdem);
  return gravou ? { ok: true } : { ok: true, repetido: true };
}

// Garante que a linha 1 tenha todas as colunas de CABECALHO (nunca apaga nem reordena nada que já exista).
// Retorna a lista de títulos da linha 1, na ordem em que estão na planilha.
function garantirCabecalho_(sh) {
  const ultimaColuna = sh.getLastColumn();
  let atual = [];
  if (sh.getLastRow() > 0 && ultimaColuna > 0) {
    atual = sh.getRange(1, 1, 1, ultimaColuna).getValues()[0].map(function (v) { return String(v).trim(); });
  }
  const linha1Vazia = atual.every(function (v) { return v === ''; });

  let criouAgora = false;
  if (linha1Vazia) {
    sh.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]);
    atual = CABECALHO.slice();
    criouAgora = true;
  } else {
    while (atual.length && atual[atual.length - 1] === '') atual.pop();
    const faltando = CABECALHO.filter(function (h) { return atual.indexOf(h) === -1; });
    if (faltando.length) {
      sh.getRange(1, atual.length + 1, 1, faltando.length).setValues([faltando]);
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

// Retorna false (sem gravar) se chaveIdem já foi usada: protege contra envio repetido (toque duplo, nova tentativa).
function gravar_(p, chaveIdem) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = chaveIdem ? CacheService.getScriptCache() : null;
    if (cache && cache.get(chaveIdem)) return false;

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
      'Captador': limpo_(p.captador)
    };

    // Monta a linha na ordem real das colunas da planilha.
    const linha = cabecalho.map(function (h) {
      return Object.prototype.hasOwnProperty.call(valores, h) ? valores[h] : '';
    });
    sh.appendRow(linha);
    const iData = cabecalho.indexOf('Data/Hora');
    if (iData !== -1) sh.getRange(sh.getLastRow(), iData + 1).setNumberFormat(FORMATO_DATA);
    if (cache) cache.put(chaveIdem, '1', 900);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// ADMINISTRAÇÃO (menu "Captadores" na planilha)
// =====================================================================

function adicionarCaptador_(login, nome, senha) {
  login = String(login || '').trim().toLowerCase();
  nome = String(nome || '').trim();
  senha = String(senha || '');
  if (!/^[a-z0-9._-]{3,30}$/.test(login)) throw new Error('Login inválido: use de 3 a 30 letras minúsculas, números, ponto, hífen ou sublinhado, sem espaços nem acentos.');
  if (nome.length < 2 || nome.length > 60) throw new Error('Nome inválido: use de 2 a 60 caracteres.');
  if (senha.length < 6 || senha.length > 100) throw new Error('A senha precisa ter de 6 a 100 caracteres.');
  const sh = abaCaptadores_(true);
  const salt = Utilities.getUuid();
  const hash = hashSenha_(salt, senha);
  const existente = buscarCaptador_(login);
  if (existente) {
    sh.getRange(existente.linha, 2, 1, 4).setValues([[nome, hash, salt, 'SIM']]);
    return 'atualizado';
  }
  sh.appendRow([login, nome, hash, salt, 'SIM']);
  return 'criado';
}
function desativarCaptador_(login) {
  const u = buscarCaptador_(String(login || '').trim().toLowerCase());
  if (!u) return false;
  abaCaptadores_(false).getRange(u.linha, 5).setValue('NAO');
  return true;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Captadores')
    .addItem('Adicionar captador ou trocar senha', 'menuAdicionarCaptador')
    .addItem('Desativar captador', 'menuDesativarCaptador')
    .addToUi();
}
function pedir_(titulo, mensagem) {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(titulo, mensagem, ui.ButtonSet.OK_CANCEL);
  return r.getSelectedButton() === ui.Button.OK ? r.getResponseText() : null;
}
function menuAdicionarCaptador() {
  const ui = SpreadsheetApp.getUi();
  const nome = pedir_('Captador (1 de 3)', 'Nome que vai aparecer na coluna Captador (ex.: Rafael B):');
  if (nome === null) return;
  const login = pedir_('Captador (2 de 3)', 'Login para entrar na página (sem espaços e sem acentos, ex.: rafael):');
  if (login === null) return;
  const senha = pedir_('Captador (3 de 3)', 'Senha (mínimo 6 caracteres). Combine com o captador por um canal privado:');
  if (senha === null) return;
  try {
    const r = adicionarCaptador_(login, nome, senha);
    ui.alert('Captador ' + r, 'Login: ' + login.trim().toLowerCase() + '\nNome na planilha: ' + nome.trim(), ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Não foi possível salvar', e.message, ui.ButtonSet.OK);
  }
}
function menuDesativarCaptador() {
  const ui = SpreadsheetApp.getUi();
  const login = pedir_('Desativar captador', 'Login do captador que não deve mais entrar:');
  if (login === null) return;
  ui.alert(desativarCaptador_(login) ? 'Captador desativado. Ele não consegue mais entrar nem cadastrar.' : 'Não encontrei esse login.');
}

// Rode UMA vez no editor (lista de funções > prepararLogin > Executar) logo depois de colar este código:
// o Google pede as permissões, e a função prepara a planilha (cabeçalho, fuso de Brasília, aba Captadores).
// Faça isso ANTES de criar a implantação.
function prepararLogin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSpreadsheetTimeZone() !== FUSO) ss.setSpreadsheetTimeZone(FUSO);
  garantirCabecalho_(ss.getSheets()[0]);
  segredo_();
  abaCaptadores_(true);
  Logger.log('Pronto: cabeçalho, fuso de Brasília, aba "Captadores" e sessão configurados.');
}
