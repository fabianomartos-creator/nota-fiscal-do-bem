/**
 * Nota Fiscal do Bem — recebe os cadastros da landing page e grava na planilha.
 *
 * Onde usar: Planilha "Inscrições - Nota Fiscal do Bem" > Extensões > Apps Script
 * (arquivo Código.gs). Depois de colar, é preciso publicar uma NOVA VERSÃO da
 * implantação (Implantar > Gerenciar implantações > editar > Nova versão).
 *
 * Este arquivo no GitHub é só a cópia versionada: ele NÃO é implantado sozinho.
 *
 * Ordem para atualizar: cole este código > Salvar > Executar "prepararLogin" (autoriza e cria a aba
 * Captadores) > Implantar > Gerenciar implantações > lápis > Nova versão > Implantar.
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

// Abrir o link /exec no navegador mostra este aviso (serve só para conferir que o serviço está no ar).
function doGet() {
  return json_({ ok: true, servico: 'Nota Fiscal do Bem', exigeLogin: exigirLogin_() });
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
  // Envio no formato antigo (página sem login). Só é aceito enquanto o login NÃO for exigido.
  if (exigirLogin_()) return { ok: false, erro: 'login_obrigatorio' };
  gravar_(p);
  return { ok: true };
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
    if (cache) cache.put(chaveIdem, '1', 900);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// LOGIN DOS CAPTADORES
// - Cada captador tem login e senha. As senhas ficam na aba "Captadores" SÓ como hash (nunca em texto).
// - Quem entra recebe uma sessão assinada (vale SESSAO_HORAS horas). O nome gravado na coluna
//   Captador vem sempre da aba "Captadores", nunca do que a página mandar.
// - Depois de MAX_FALHAS senhas erradas, o login fica bloqueado por BLOQUEIO_SEGUNDOS.
// - Só passa a ser OBRIGATÓRIO quando você ligar no menu: Captadores > Exigir login: LIGAR.
// =====================================================================
const ABA_CAPTADORES = 'Captadores';
const CAB_CAPTADORES = ['Login', 'Nome', 'Hash da senha', 'Salt', 'Ativo'];
const MAX_FALHAS = 5;
const BLOQUEIO_SEGUNDOS = 900;
const SESSAO_HORAS = 12;
const PROP_EXIGIR = 'EXIGIR_LOGIN';      // 'sim' ou 'nao'
const PROP_SEGREDO = 'SEGREDO_SESSAO';

function exigirLogin_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_EXIGIR) === 'sim';
}

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
    sh = ss.insertSheet(ABA_CAPTADORES, ss.getSheets().length); // sempre no fim: a 1ª aba continua sendo a das inscrições
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
    utm_source: p.utm_source, utm_medium: p.utm_medium, utm_campaign: p.utm_campaign, pagina: p.pagina,
    captador: u.nome            // vem da aba Captadores, não da página
  }, chaveIdem);
  return gravou ? { ok: true } : { ok: true, repetido: true };
}

// ---- administração (menu "Captadores" na planilha) ----
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
function existeCaptadorAtivo_() {
  const sh = abaCaptadores_(false);
  if (!sh || sh.getLastRow() < 2) return false;
  return sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues().some(function (l) {
    return String(l[0]).trim().toUpperCase() === 'SIM';
  });
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Captadores')
    .addItem('Adicionar captador ou trocar senha', 'menuAdicionarCaptador')
    .addItem('Desativar captador', 'menuDesativarCaptador')
    .addSeparator()
    .addItem('Exigir login na página: LIGAR', 'menuExigirLigar')
    .addItem('Exigir login na página: DESLIGAR', 'menuExigirDesligar')
    .addItem('Ver se o login está sendo exigido', 'menuStatusLogin')
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
function menuExigirLigar() {
  const ui = SpreadsheetApp.getUi();
  if (!existeCaptadorAtivo_()) {
    ui.alert('Antes de exigir o login, cadastre pelo menos um captador (menu Captadores > Adicionar captador).');
    return;
  }
  PropertiesService.getScriptProperties().setProperty(PROP_EXIGIR, 'sim');
  ui.alert('Login EXIGIDO. A partir de agora só quem entra com login e senha consegue cadastrar.');
}
function menuExigirDesligar() {
  PropertiesService.getScriptProperties().setProperty(PROP_EXIGIR, 'nao');
  SpreadsheetApp.getUi().alert('Login NÃO exigido. A página antiga, sem login, volta a ser aceita.');
}
function menuStatusLogin() {
  SpreadsheetApp.getUi().alert(exigirLogin_() ? 'O login está sendo EXIGIDO.' : 'O login NÃO está sendo exigido (a página antiga ainda é aceita).');
}
// Rode UMA vez no editor (lista de funções > prepararLogin > Executar) logo depois de colar este código:
// o Google pede as permissões novas e cria a aba "Captadores". Faça isso ANTES de publicar a nova versão.
function prepararLogin() {
  segredo_();
  abaCaptadores_(true);
  Logger.log('Pronto: aba "Captadores" criada e sessão configurada. Agora publique a Nova versão.');
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
