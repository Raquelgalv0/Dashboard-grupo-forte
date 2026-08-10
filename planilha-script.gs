/**
 * Radar editorial, Grupo FORTE
 * Porteiro da planilha: qualquer pessoa lê o painel, mas só quem tem a senha
 * inclui ou remove chamadas. A senha é conferida aqui, no servidor do Google,
 * então não dá para burlar pelo navegador.
 */

const COLUNAS = ['name', 'title', 'status', 'scope', 'theme', 'deadline',
                 'fit', 'tags', 'description', 'authorship', 'cost', 'url'];

/** Rode UMA VEZ para definir a senha do grupo (troque o valor abaixo). */
function definirSenha() {
  PropertiesService.getScriptProperties().setProperty('SENHA', 'TROQUE-ESTA-SENHA');
}

/** Usa a aba "chamadas" se existir; senão, a primeira aba da planilha. */
function abrirAba_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName('chamadas') || planilha.getSheets()[0];
  if (aba.getLastRow() === 0) aba.appendRow(COLUNAS);
  return aba;
}

function lerTudo_() {
  const aba = abrirAba_();
  const linhas = aba.getDataRange().getValues();
  if (linhas.length < 2) return [];

  const cabecalho = linhas[0].map(c => String(c).trim());
  const colunaUrl = cabecalho.indexOf('url');
  if (colunaUrl === -1) return [];

  return linhas.slice(1)
    .filter(linha => String(linha[colunaUrl] || '').trim())
    .map(linha => {
      const chamada = {};
      COLUNAS.forEach(coluna => {
        const i = cabecalho.indexOf(coluna);
        chamada[coluna] = i === -1 ? '' : String(linha[i] == null ? '' : linha[i]);
      });
      chamada.tags = chamada.tags
        ? chamada.tags.split(',').map(t => t.trim()).filter(Boolean)
        : [];
      return chamada;
    });
}

function responder_(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Leitura pública: o dashboard chama isto para montar os cards. */
function doGet() {
  try {
    return responder_({ ok: true, data: lerTudo_() });
  } catch (erro) {
    return responder_({ ok: false, error: String(erro) });
  }
}

/** Escrita protegida: incluir ou remover chamada. */
function doPost(e) {
  try {
    const corpo = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    const senhaConfigurada = PropertiesService.getScriptProperties().getProperty('SENHA');
    if (!senhaConfigurada || senhaConfigurada === 'TROQUE-ESTA-SENHA') {
      return responder_({ ok: false, error: 'A senha ainda não foi definida no script.' });
    }
    if (String(corpo.senha || '') !== senhaConfigurada) {
      return responder_({ ok: false, error: 'Senha incorreta.' });
    }

    const aba = abrirAba_();

    if (corpo.acao === 'incluir') {
      const nova = corpo.chamada || {};
      if (!String(nova.name || '').trim() || !String(nova.url || '').trim()) {
        return responder_({ ok: false, error: 'Informe ao menos o nome da revista e o link.' });
      }
      if (lerTudo_().some(c => c.url === String(nova.url).trim())) {
        return responder_({ ok: false, error: 'Essa chamada já está no painel.' });
      }
      const cabecalho = aba.getDataRange().getValues()[0].map(c => String(c).trim());
      aba.appendRow(cabecalho.map(coluna => {
        if (coluna === 'tags') {
          return Array.isArray(nova.tags) ? nova.tags.join(', ') : (nova.tags || '');
        }
        return nova[coluna] || '';
      }));
      return responder_({ ok: true, data: lerTudo_() });
    }

    if (corpo.acao === 'remover') {
      const alvo = String(corpo.url || '').trim();
      if (!alvo) return responder_({ ok: false, error: 'Informe o link da chamada a remover.' });

      const linhas = aba.getDataRange().getValues();
      const colunaUrl = linhas[0].map(c => String(c).trim()).indexOf('url');
      let removidas = 0;
      for (let i = linhas.length - 1; i >= 1; i--) {
        if (String(linhas[i][colunaUrl]).trim() === alvo) {
          aba.deleteRow(i + 1);
          removidas++;
        }
      }
      if (!removidas) return responder_({ ok: false, error: 'Não encontrei nenhuma chamada com esse link.' });
      return responder_({ ok: true, data: lerTudo_() });
    }

    return responder_({ ok: false, error: 'Ação desconhecida.' });
  } catch (erro) {
    return responder_({ ok: false, error: String(erro) });
  }
}
