/**
 * IDIOMA IA — servidor (Google Apps Script, vinculado a uma planilha Google)
 *
 * Cole este arquivo inteiro em Code.gs, rode a função `instalar` uma vez e publique
 * como App da Web (Executar como: Eu / Quem pode acessar: Qualquer pessoa).
 *
 * O que este servidor faz:
 *   - cadastro e login por e-mail e senha, só para quem tem o código de convite;
 *   - guarda o progresso de cada pessoa na planilha;
 *   - ouve a voz do aluno nos desafios de fala e avalia a pronúncia;
 *   - chama o Gemini com a chave guardada em Propriedades do script (GEMINI_API_KEY),
 *     que nunca vai para o celular de ninguém.
 */

/* ============================== configuração ============================== */
const ABAS = {
  CONFIG:   ['chave', 'valor', 'para que serve'],
  USERS:    ['user_id', 'name', 'email', 'pass_hash', 'salt', 'target_language', 'level', 'xp', 'streak', 'last_activity', 'daily_minutes', 'created_at'],
  SESSIONS: ['token_hash', 'user_id', 'created_at', 'expires_at'],
  STATE:    ['user_id', 'part', 'json', 'updated_at'],
  USAGE:    ['day', 'user_id', 'calls', 'tokens_in', 'tokens_out']
};
const PADRAO = {
  modelo: 'gemini-3.5-flash-lite', // modelo do Gemini usado pelo professor
  modelo_voz: 'gemini-3.8-flash-tts', // modelo do Gemini que lê em voz alta
  limite_diario: 300,              // chamadas de IA por pessoa, por dia
  max_pessoas: 12,                 // quantas contas podem existir
  dias_sessao: 60                  // dias até pedir login de novo
};
const PEDACO = 40000;   // uma célula guarda até 50 mil caracteres; o progresso é dividido em partes
const MAX_ESTADO = 400000;
const MAX_PROMPT = 60000;

/* ============================== instalação ============================== */
/** Rode uma vez pelo editor. Cria as abas e o código de convite. Pode rodar de novo sem perder nada. */
function instalar() {
  Object.keys(ABAS).forEach(function (n) { aba_(n); });
  const cfg = config_();
  const sh = aba_('CONFIG');
  if (!cfg.convite) sh.appendRow(['convite', novoConvite_(), 'Código que a pessoa digita para criar conta. Troque quando quiser.']);
  if (!cfg.modelo) sh.appendRow(['modelo', PADRAO.modelo, 'Modelo do Gemini usado pelo professor.']);
  if (!cfg.limite_diario) sh.appendRow(['limite_diario', PADRAO.limite_diario, 'Chamadas de IA por pessoa, por dia.']);
  if (!cfg.max_pessoas) sh.appendRow(['max_pessoas', PADRAO.max_pessoas, 'Quantas contas podem existir.']);
  if (!cfg.modelo_voz) sh.appendRow(['modelo_voz', PADRAO.modelo_voz, 'Modelo do Gemini que lê as frases em voz alta.']);
  const temChave = !!prop_('GEMINI_API_KEY');
  Logger.log('Abas prontas. Código de convite: ' + config_().convite);
  Logger.log(temChave ? 'Chave do Gemini encontrada.' : 'FALTA a chave: Configurações do projeto > Propriedades do script > GEMINI_API_KEY.');
  return { convite: config_().convite, chave: temChave };
}
/** Rode pelo editor para conferir se a chave do Gemini funciona. */
function testarGemini() {
  const r = gemini_([{ role: 'user', parts: [{ text: 'Responda somente com o JSON {"ok": true}' }] }]);
  Logger.log('Gemini respondeu: ' + r.text + ' (' + r.tokensIn + ' tokens de entrada, ' + r.tokensOut + ' de saída)');
  return r.text;
}
function novoConvite_() {
  const letras = 'ABCDEFGHJKLMNPQRSTUVWXYZ'; // só letras: a planilha não confunde com número
  let s = '';
  const u = Utilities.getUuid().replace(/-/g, '');
  for (let i = 0; i < 6; i++) s += letras.charAt(parseInt(u.substr(i * 2, 2), 16) % letras.length);
  return s;
}

/* ============================== entrada ============================== */
function doGet() {
  return saida_({ ok: true, data: { app: 'Idioma IA', versao: 1, chave: !!prop_('GEMINI_API_KEY') } });
}
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return saida_(falha_('bad_request', 'Pedido inválido.')); }
  try {
    const a = String(req.action || '');
    if (a === 'register') return saida_(ok_(registrar_(req)));
    if (a === 'login') return saida_(ok_(entrar_(req)));
    const user = autenticar_(req.token);
    if (a === 'logout') return saida_(ok_(sair_(req.token)));
    if (a === 'getState') return saida_(ok_(lerEstado_(user)));
    if (a === 'saveState') return saida_(ok_(gravarEstado_(user, req)));
    if (a === 'ai') return saida_(ok_(ia_(user, req)));
    if (a === 'tts') return saida_(ok_(voz_(user, req)));
    if (a === 'ouvir') return saida_(ok_(ouvir_(user, req)));
    return saida_(falha_('bad_request', 'Ação desconhecida.'));
  } catch (err) {
    if (err && err.code) return saida_(falha_(err.code, err.message));
    console.error(err);
    return saida_(falha_('server', 'Erro no servidor: ' + (err && err.message ? err.message : err)));
  }
}
function ok_(data) { return { ok: true, data: data }; }
function falha_(code, message) { return { ok: false, error: { code: code, message: message } }; }
function saida_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function erro_(code, message) { return { code: code, message: message }; }

/* ============================== planilha ============================== */
function aba_(nome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh) { sh = ss.insertSheet(nome); sh.appendRow(ABAS[nome]); sh.setFrozenRows(1); }
  return sh;
}
/** Linhas da aba, sem o cabeçalho. A linha i está na linha i + 2 da planilha. */
function linhas_(nome) { return aba_(nome).getDataRange().getValues().slice(1); }
function config_() {
  const o = {};
  linhas_('CONFIG').forEach(function (r) { if (r[0] !== '') o[String(r[0]).trim()] = r[1]; });
  return o;
}
function numCfg_(cfg, k) { const n = Number(cfg[k]); return n > 0 ? n : PADRAO[k]; }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function comTrava_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function hoje_() { return Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')); }

/* ============================== contas ============================== */
function hash_(texto) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  let hex = '';
  for (let i = 0; i < bytes.length; i++) { const b = (bytes[i] + 256) % 256; hex += (b < 16 ? '0' : '') + b.toString(16); }
  return hex;
}
/** Hash da senha com sal, repetido para encarecer tentativas. O prefixo impede a planilha de ler como número. */
function hashSenha_(senha, sal) {
  let h = sal + ':' + senha;
  for (let i = 0; i < 300; i++) h = hash_(sal + h);
  return 'h:' + h;
}
function email_(v) { return String(v || '').trim().toLowerCase(); }
function registrar_(req) {
  const nome = String(req.name || '').trim().slice(0, 40);
  const email = email_(req.email);
  const senha = String(req.password || '');
  if (!nome) throw erro_('invalid', 'Informe seu nome.');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw erro_('invalid', 'Confira o e-mail.');
  if (senha.length < 6) throw erro_('weak', 'A senha precisa de pelo menos 6 caracteres.');
  const cfg = config_();
  const convite = String(cfg.convite || '').trim().toUpperCase();
  if (!convite || String(req.invite || '').trim().toUpperCase() !== convite) throw erro_('invite', 'Código de convite incorreto.');
  return comTrava_(function () {
    const users = linhas_('USERS');
    if (users.some(function (r) { return email_(r[2]) === email; })) throw erro_('exists', 'Já existe uma conta com esse e-mail. Use Entrar.');
    if (users.length >= numCfg_(cfg, 'max_pessoas')) throw erro_('full', 'O limite de contas foi atingido. Fale com quem administra o app.');
    const id = 'u_' + Utilities.getUuid().replace(/-/g, '');
    const sal = Utilities.getUuid();
    aba_('USERS').appendRow([id, nome, email, hashSenha_(senha, sal), sal, '', '', 0, 0, '', '', Date.now()]);
    return { token: novaSessao_(id), user: { id: id, name: nome, email: email } };
  });
}
function entrar_(req) {
  const email = email_(req.email);
  const senha = String(req.password || '');
  const cache = CacheService.getScriptCache();
  const chave = 'tent:' + email;
  const tent = Number(cache.get(chave) || 0);
  if (tent >= 8) throw erro_('rate_limited', 'Muitas tentativas. Espere 10 minutos e tente de novo.');
  const u = linhas_('USERS').filter(function (r) { return email_(r[2]) === email; })[0];
  if (!u || hashSenha_(senha, String(u[4])) !== String(u[3])) {
    cache.put(chave, String(tent + 1), 600);
    throw erro_('credentials', 'E-mail ou senha incorretos.');
  }
  return comTrava_(function () {
    limparSessoes_();
    return { token: novaSessao_(String(u[0])), user: { id: String(u[0]), name: String(u[1]), email: email } };
  });
}
function novaSessao_(userId) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  const dias = PADRAO.dias_sessao;
  aba_('SESSIONS').appendRow(['h:' + hash_(token), userId, Date.now(), Date.now() + dias * 864e5]);
  CacheService.getScriptCache().put('sess:' + hash_(token), userId, 21600);
  return token;
}
function limparSessoes_() {
  const sh = aba_('SESSIONS');
  const v = linhas_('SESSIONS');
  for (let i = v.length - 1; i >= 0; i--) if (Number(v[i][3]) < Date.now()) sh.deleteRow(i + 2);
}
function autenticar_(token) {
  token = String(token || '');
  if (token.length < 32) throw erro_('auth', 'Entre de novo.');
  const th = hash_(token);
  const cache = CacheService.getScriptCache();
  let id = cache.get('sess:' + th);
  if (!id) {
    const s = linhas_('SESSIONS').filter(function (r) { return String(r[0]) === 'h:' + th && Number(r[3]) > Date.now(); })[0];
    if (!s) throw erro_('auth', 'Sua sessão terminou. Entre de novo.');
    id = String(s[1]);
    cache.put('sess:' + th, id, 21600);
  }
  return { id: id };
}
function sair_(token) {
  const th = hash_(String(token));
  CacheService.getScriptCache().remove('sess:' + th);
  comTrava_(function () {
    const sh = aba_('SESSIONS');
    const v = linhas_('SESSIONS');
    for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === 'h:' + th) sh.deleteRow(i + 2);
  });
  return { saiu: true };
}

/* ============================== progresso ============================== */
function lerEstado_(user) {
  const partes = linhas_('STATE').filter(function (r) { return String(r[0]) === user.id; })
    .sort(function (a, b) { return Number(a[1]) - Number(b[1]); });
  if (!partes.length) return { state: null, updatedAt: 0 };
  return { state: partes.map(function (r) { return String(r[2]).slice(1); }).join(''), updatedAt: Number(partes[0][3]) || 0 };
}
function gravarEstado_(user, req) {
  const txt = String(req.state || '');
  if (!txt || txt.length > MAX_ESTADO) throw erro_('invalid', 'Progresso grande demais para gravar.');
  let st;
  try { st = JSON.parse(txt); } catch (err) { throw erro_('invalid', 'Progresso ilegível.'); }
  if (!st || typeof st !== 'object') throw erro_('invalid', 'Progresso ilegível.');
  const quando = Number(req.updatedAt) || Date.now();
  return comTrava_(function () {
    const sh = aba_('STATE');
    const v = linhas_('STATE');
    for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === user.id) sh.deleteRow(i + 2);
    for (let p = 0, k = 0; p < txt.length; p += PEDACO, k++) sh.appendRow([user.id, k, '#' + txt.substr(p, PEDACO), quando]); // '#' impede a planilha de ler o pedaço como fórmula ou número
    // resumo legível na aba USERS
    const us = aba_('USERS');
    const u = linhas_('USERS');
    for (let i = 0; i < u.length; i++) if (String(u[i][0]) === user.id) {
      const pf = st.profile || {};
      us.getRange(i + 2, 6, 1, 6).setValues([[String(pf.target || ''), String(st.level || ''), Number(st.xp) || 0,
        Number(st.streak && st.streak.cur) || 0, String((st.streak && st.streak.last) || ''), Number(pf.minutes) || '']]);
      break;
    }
    return { updatedAt: quando };
  });
}

/* ============================== professor (Gemini) ============================== */
function ia_(user, req) {
  // aceita um texto só, ou turnos [{role: 'user' | 'assistant', content}]
  let turnos = req.input;
  if (typeof turnos === 'string') turnos = [{ role: 'user', content: turnos }];
  if (!Array.isArray(turnos) || !turnos.length) throw erro_('invalid', 'Pedido vazio.');
  const contents = [];
  let total = 0;
  turnos.forEach(function (t) {
    const texto = String((t && t.content) || '');
    if (!texto) return;
    total += texto.length;
    const role = t.role === 'assistant' ? 'model' : 'user';
    const ult = contents[contents.length - 1];
    if (ult && ult.role === role) ult.parts[0].text += '\n\n' + texto; // o Gemini espera papéis alternados
    else contents.push({ role: role, parts: [{ text: texto }] });
  });
  if (!contents.length || contents[contents.length - 1].role !== 'user') throw erro_('invalid', 'O pedido precisa terminar com a fala do aluno.');
  if (total > MAX_PROMPT) throw erro_('invalid', 'Pedido grande demais.');

  checarLimite_(user);
  const r = gemini_(contents, String(config_().modelo || PADRAO.modelo));
  contarUso_(user, r.tokensIn, r.tokensOut);
  return r;
}
function checarLimite_(user) {
  const dia = hoje_();
  const limite = numCfg_(config_(), 'limite_diario');
  const usado = linhas_('USAGE').filter(function (r) { return Number(r[0]) === dia && String(r[1]) === user.id; })[0];
  if (usado && Number(usado[2]) >= limite) throw erro_('limit', 'Você chegou ao limite de hoje com o professor. Amanhã tem mais.');
}
function contarUso_(user, tokensIn, tokensOut) {
  const dia = hoje_();
  comTrava_(function () {
    const sh = aba_('USAGE');
    const v = linhas_('USAGE');
    for (let i = v.length - 1; i >= 0; i--) if (Number(v[i][0]) === dia && String(v[i][1]) === user.id) {
      sh.getRange(i + 2, 3, 1, 3).setValues([[Number(v[i][2]) + 1, Number(v[i][3]) + tokensIn, Number(v[i][4]) + tokensOut]]);
      return;
    }
    sh.appendRow([dia, user.id, 1, tokensIn, tokensOut]);
  });
}
/** Ouve a voz do aluno (WAV em base64) e avalia com o Gemini, seguindo o pedido do app. */
function ouvir_(user, req) {
  const audio = String(req.audio || '');
  if (!audio || audio.length > 4000000 || !/^[A-Za-z0-9+/=]+$/.test(audio.slice(0, 200))) throw erro_('invalid', 'Áudio vazio ou longo demais.');
  const pedido = String(req.prompt || '');
  if (!pedido) throw erro_('invalid', 'Pedido vazio.');
  if (pedido.length > MAX_PROMPT) throw erro_('invalid', 'Pedido grande demais.');
  const mime = /^audio\/[a-z0-9.+-]+$/.test(String(req.mime || '')) ? String(req.mime) : 'audio/wav';
  checarLimite_(user);
  const r = gemini_([{ role: 'user', parts: [{ text: pedido }, { inlineData: { mimeType: mime, data: audio } }] }]);
  contarUso_(user, r.tokensIn, r.tokensOut);
  return r;
}
/** Lê uma frase em voz alta com uma voz do Gemini. Devolve o áudio em base64 (WAV ou PCM). */
function voz_(user, req) {
  const texto = String(req.text || '').trim().slice(0, 400);
  if (!texto) throw erro_('invalid', 'Nada para falar.');
  const voz = /^[A-Za-z]{2,24}$/.test(String(req.voice || '')) ? String(req.voice) : 'Kore';
  checarLimite_(user);
  const modelo = String(config_().modelo_voz || PADRAO.modelo_voz);
  const estilo = req.slow ? 'speaking slowly and clearly, like a patient language teacher' : 'speaking clearly, like a friendly language teacher';
  // Primeiro o formato atual da documentação; se o modelo recusar, o formato anterior.
  const j = postGemini_(modelo, [
    { contents: [{ role: 'user', parts: [{ text: texto, speech_metadata: { style: estilo } }] }], generationConfig: { responseModalities: ['AUDIO'], speechConfig: { voiceConfig: { voice: voz } } } },
    { contents: [{ role: 'user', parts: [{ text: texto }] }], generationConfig: { responseModalities: ['AUDIO'], speechConfig: { voiceConfig: { prebuiltVoiceConfig: { voiceName: voz } } } } }
  ], 'modelo_voz');
  const partes = (j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts) || [];
  const p = partes.filter(function (x) { return x && x.inlineData && x.inlineData.data; })[0];
  if (!p) throw erro_('refused', 'O Gemini não devolveu o áudio.');
  const um = j.usageMetadata || {};
  contarUso_(user, Number(um.promptTokenCount) || 0, Number(um.candidatesTokenCount) || 0);
  return { audio: p.inlineData.data, mime: String(p.inlineData.mimeType || 'audio/wav') };
}
/** Envia ao Gemini, tentando cada corpo em ordem enquanto o modelo recusar o formato (erro 400). */
function postGemini_(modelo, corpos, linhaConfig) {
  const chave = prop_('GEMINI_API_KEY');
  if (!chave) throw erro_('config', 'Falta a chave do Gemini no servidor (Propriedades do script > GEMINI_API_KEY).');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelo) + ':generateContent';
  let ultimo = '';
  for (let i = 0; i < corpos.length; i++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': chave },
      payload: JSON.stringify(corpos[i]), muteHttpExceptions: true
    });
    const status = res.getResponseCode();
    const txt = res.getContentText();
    if (status === 200) {
      try { return JSON.parse(txt); } catch (err) { throw erro_('upstream_error', 'Resposta ilegível do Gemini.'); }
    }
    ultimo = status + ' ' + txt.slice(0, 300);
    if (status === 429) throw erro_('rate_limited', 'O Gemini está no limite de uso agora. Tente de novo em instantes.');
    if (status === 401 || status === 403 || /API key/i.test(txt)) throw erro_('config', 'O Gemini recusou a chave do servidor. Confira GEMINI_API_KEY.');
    if (status === 404) throw erro_('config', 'O modelo "' + modelo + '" não foi encontrado. Confira a linha "' + linhaConfig + '" na aba CONFIG.');
    if (status !== 400) throw erro_('upstream_error', 'O Gemini falhou (' + status + '). Tente de novo.');
  }
  console.error('Gemini 400: ' + ultimo);
  throw erro_('upstream_error', 'O Gemini recusou o pedido. Veja o registro de execuções do Apps Script.');
}
function gemini_(contents, modelo) {
  modelo = modelo || String(config_().modelo || PADRAO.modelo);
  // Tenta a configuração completa; se o modelo recusar algum campo, tenta de novo com menos.
  const j = postGemini_(modelo, [
    { contents: contents, generationConfig: { temperature: 0.8, responseMimeType: 'application/json', thinkingConfig: { thinkingLevel: 'low' } } },
    { contents: contents, generationConfig: { temperature: 0.8, responseMimeType: 'application/json' } },
    { contents: contents }
  ], 'modelo');
  const cand = j.candidates && j.candidates[0];
  const partes = (cand && cand.content && cand.content.parts) || [];
  const texto = partes.filter(function (p) { return p && !p.thought && typeof p.text === 'string'; }).map(function (p) { return p.text; }).join('');
  if (!texto.trim()) throw erro_('refused', 'O professor não respondeu a esse pedido.');
  const um = j.usageMetadata || {};
  return { text: texto, tokensIn: Number(um.promptTokenCount) || 0, tokensOut: (Number(um.candidatesTokenCount) || 0) + (Number(um.thoughtsTokenCount) || 0) };
}
