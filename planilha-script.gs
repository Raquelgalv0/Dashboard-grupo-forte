/**
 * Radar editorial, Grupo FORTE
 * Porteiro da planilha: qualquer pessoa lê o painel, mas só quem tem a senha
 * inclui ou remove chamadas. A senha é conferida aqui, no servidor do Google,
 * então não dá para burlar pelo navegador.
 */

const ABA = 'chamadas';
const COLUNAS = ['name', 'title', 'status', 'scope', 'theme', 'deadline',
                 'fit', 'tags', 'description', 'authorship', 'cost', 'url'];

/** Rode UMA VEZ para definir a senha do grupo (troque o valor abaixo). */
function definirSenha() {
  PropertiesService.getScriptProperties().setProperty('SENHA', 'TROQUE-ESTA-SENHA');
}

function abrirAba_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(ABA);
  if (!aba) {
    aba = planilha.insertSheet(ABA);
    aba.appendRow(COLUNAS);
    aba.setFrozenRows(1);
  }
  return aba;
}

function lerTudo_() {
  const aba = abrirAba_();
  const linhas = aba.getDataRange().getValues();
  if (linhas.length < 2) return [];

  const cabecalho = linhas[0];
  const colunaUrl = cabecalho.indexOf('url');

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
      aba.appendRow(COLUNAS.map(coluna => {
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
      const colunaUrl = linhas[0].indexOf('url');
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

/** Rode UMA VEZ para levar para a planilha as chamadas que já estão no painel. */
function carregarChamadasAtuais() {
  const atuais = [
    ['Pensares em Revista', 'Dossiê 38: Emoção, formação de professores e Inteligência Artificial', 'Chamada aberta', 'Nacional', 'Emoções e IA', '30/09/2026', 'Muito alta', 'emoções, formação docente, IA generativa, ética', 'Dossiê sobre emoções, cognição, linguagem, formação inicial e continuada, práticas interacionais com sistemas de IA e justiça sociotécnica.', 'Conferir diretrizes da revista', 'Acesso aberto, verificar taxas', 'https://www.e-publicacoes.uerj.br/index.php/pensaresemrevista'],
    ['Educação em Análise', 'A formação de professores na cibercultura', 'Chamada aberta', 'Nacional', 'Formação docente', '30/08/2026', 'Alta', 'cibercultura, IA, ética, formação', 'Chamada sobre formação docente articulada a dimensões sociotécnicas, éticas, estéticas, políticas e culturais.', 'Estudantes de pós-graduação são convidados', 'Verificar nas diretrizes', 'https://ojs.uel.br/revistas/uel/index.php/educanalise'],
    ['Learning, Media and Technology', 'Towards a Critical Approach to AI in Education', 'Chamada aberta', 'Internacional', 'IA crítica e ética', 'Resumo: 31/08/2026', 'Muito alta', 'critical AI, power, education, ethics', 'Special issue internacional voltado a abordagens críticas da IA, controvérsias, relações sociomateriais, conhecimento e poder.', 'Artigo em inglês', 'Verificar opção de publicação', 'https://think.taylorandfrancis.com/special_issues/towards-a-critical-approach-to-ai-in-education/'],
    ['Trabalhos em Linguística Aplicada', 'Submissões em fluxo contínuo', 'Fluxo contínuo', 'Nacional', 'Linguística Aplicada', 'Fluxo contínuo', 'Muito alta', 'Linguística Aplicada, tecnologias, educação linguística', 'Periódico de Linguística Aplicada com submissões em fluxo contínuo e interesse em linguagem, educação linguística, tecnologias e vida social.', 'Ao menos uma pessoa doutora ou doutoranda', 'Verificar nas diretrizes', 'https://periodicos.sbu.unicamp.br/ojs/index.php/tla/about/submissions'],
    ['Anthesis (UFAC)', 'Dossiê: Insurgência Nada Artificial, autoria, criticidade e letramentos em tempos de inteligência artificial', 'Chamada aberta', 'Nacional', 'IA crítica e ética', '29/11/2026', 'Muito alta', 'IA generativa, autoria, letramentos, ética', 'Dossiê sobre IA generativa, educação, escrita acadêmica, autoria, ética e cultura digital. Revista do Programa de Pós-Graduação em Ensino de Humanidades e Linguagens da UFAC.', 'Conferir diretrizes da revista', 'Verificar nas diretrizes', 'https://periodicos.ufac.br/index.php/anthesis']
  ];

  const aba = abrirAba_();
  if (aba.getLastRow() > 1) {
    throw new Error('A planilha já tem chamadas. Apague-as antes de rodar isto de novo.');
  }
  atuais.forEach(linha => aba.appendRow(linha));
}
