/**
 * MULTIWASHER - BACKEND GOOGLE SHEETS + DRIVE (v3)
 *
 * Novidades face à v2:
 *   - As fotos e o logótipo deixam de ir em base64 para a célula: são guardados
 *     no Drive, numa subpasta por relatório, e no Sheet ficam os URLs.
 *   - O próprio Web App passa a servir o relatório numa página partilhável:
 *         .../exec?id=MW-ARCOR-2026-01
 *     Esse link abre em qualquer equipamento, sem precisar da app.
 *   - v3.1: a página partilhável (link e ficheiro descarregado) passou a incluir
 *     o Resumo Executivo (Pontos Fortes / Pontos a Melhorar / Recomendações /
 *     Riscos & Alertas), gerado automaticamente a partir dos dados do ensaio —
 *     a mesma lógica que já existe na app (index.html).
 *
 * Funções:
 *   - doPost(e)              -> recebe o relatório, grava fotos no Drive e linhas no Sheet
 *   - doGet(e)               -> sem parâmetros: estado | ?id=XXX: página do relatório
 *                               | ?id=XXX&formato=json: dados em JSON
 *   - atualizarCabecalho()   -> acrescenta a coluna nova SEM apagar dados (correr uma vez)
 *   - setupSheetHeaders()    -> só para começar do zero: APAGA tudo
 *
 * Implementar -> Nova implementação -> Aplicação Web
 *   Executar como:   Eu
 *   Quem tem acesso: Qualquer pessoa
 */

var SHEET_NAME = "Folha1";
var MAX_CELL = 45000;
var COL_ID = 2;

// Pasta do Drive onde ficam as subpastas de cada relatório
var PASTA_RAIZ_ID = "1bCvjoSq-96_fggxWWSPhG2Fm-BczOCuE"; // Projeto ID Material de Cliente

/**
 * Quem pode ver as fotos:
 *   "publico"  -> qualquer pessoa com o link (necessário para clientes externos)
 *   "dominio"  -> só quem tiver conta @somengil.com
 *   "privado"  -> só quem já tem acesso à pasta (as imagens não aparecem a terceiros)
 */
var PARTILHA = "publico";

var HEADERS = [
  "Timestamp",
  "ID Relatório",
  "Data Ensaio",
  "Cliente",
  "Logótipo Cliente (URL)",
  "Setor / Indústria",
  "Técnico Responsável",
  "Modelo Máquina",
  "Nº Série",
  "Nº Utensílio",
  "Total Utensílios",
  "Utensílio - Tipo",
  "Utensílio - Material",
  "Utensílio - Sujidade",
  "Utensílio - Objetivo",
  "Vapor Inicial",
  "Lavagem",
  "Centr. Intermédia",
  "Enxaguamento",
  "Centr. Final",
  "Vapor Final",
  "Considerações Prévias",
  "Considerações Práticas",
  "Conclusão Final",
  "Fotos / Evidências (JSON)",
  "Vídeo URL",
  "Pasta Drive (URL)",
  "Resumo Executivo Manual (JSON)"
];

var LOGO_SOMENGIL = "https://res.cloudinary.com/dlkkjtgvy/image/upload/v1786032688/1_SOMENGIL_PNG_LOGO__p1z4co.png";

/** ============================ RECEÇÃO ============================ **/

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(60000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ ok: false, error: "Sem corpo no pedido (postData vazio)." });
    }

    var d = JSON.parse(e.postData.contents);

    // Eliminar um relatório (pedido pela app)
    if (d.acao === "apagar") {
      if (!d.id) return jsonOut({ ok: false, error: "Falta o ID do relatório a apagar." });
      var apagadasAgora = apagarLinhasDoId(getSheet(), d.id);
      var pastaLixo = mandarPastaParaLixo(d.id);
      return jsonOut({
        ok: true, acao: "apagado", id: String(d.id),
        linhas_apagadas: apagadasAgora, pasta_no_lixo: pastaLixo
      });
    }

    if (!d.id || String(d.id).trim() === "") {
      return jsonOut({ ok: false, error: "O relatório não tem ID. Preencha o ID do Relatório antes de enviar." });
    }

    var sheet = getSheet();
    var pasta = pastaDoRelatorio(d.id);

    // Logótipo do cliente
    var logoUrl = guardarImagem(pasta, "logotipo-cliente", d.clienteLogo);

    // Fotos de evidência
    var fotos = [];
    (d.fotos || []).forEach(function (f, i) {
      var desc = f && f.desc ? String(f.desc) : "";
      var url = guardarImagem(pasta, "foto-" + pad2(i + 1), f && f.url);
      if (url || desc) fotos.push({ n: i + 1, desc: desc, url: url || "" });
    });

    var utensilios = normalizarUtensilios(d);
    var total = utensilios.length;
    var fotosJson = fotos.length ? JSON.stringify(fotos) : "";
    var pastaUrl = pasta ? pasta.getUrl() : "";

    var linhas = utensilios.map(function (u, i) {
      return [
        new Date(),
        txt(d.id), txt(d.data), txt(d.cliente), txt(logoUrl), txt(d.setor),
        txt(d.tecnico), txt(d.maquina), txt(d.serialNumber),
        i + 1, total,
        txt(u.tipo), txt(u.material), txt(u.sujidade), txt(u.objetivo),
        txt(u.vaporIni), txt(u.lavagem), txt(u.centrifInter),
        txt(u.enxague), txt(u.centrifFin), txt(u.vaporFin),
        txt(d.consideracoesPrevias), txt(d.consideracoes), txt(d.conclusao),
        txt(fotosJson), txt(d.videoUrl), txt(pastaUrl),
        txt(d.resumoManual ? JSON.stringify(d.resumoManual) : "")
      ];
    });

    var apagadas = apagarLinhasDoId(sheet, d.id);
    sheet.getRange(sheet.getLastRow() + 1, 1, linhas.length, HEADERS.length).setValues(linhas);

    return jsonOut({
      ok: true,
      id: txt(d.id),
      acao: apagadas > 0 ? "substituído" : "inserido",
      linhas: linhas.length,
      utensilios: total,
      fotos_guardadas: fotos.filter(function (f) { return f.url; }).length,
      pasta: pastaUrl,
      link: linkDoRelatorio(d.id)
    });

  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** ============================ LEITURA ============================ **/

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var id = p.id || null;

  // Lista de relatórios (para o histórico da app)
  if (p.acao === "lista") {
    return jsonOut({ ok: true, relatorios: listarRelatorios() });
  }

  // Todos os relatórios completos (para importar de uma vez para a app)
  if (p.acao === "exportar") {
    var ids = listarRelatorios().map(function (r) { return r.id; });
    var todos = ids.map(function (x) { return lerRelatorio(x); }).filter(function (r) { return r; });
    return jsonOut({ ok: true, total: todos.length, relatorios: todos });
  }

  if (!id) {
    var sheet = getSheet();
    return jsonOut({
      ok: true,
      mensagem: "Web App MultiWasher ativo (v3 - fotos no Drive, relatório partilhável).",
      folha: SHEET_NAME,
      colunas: HEADERS.length,
      linhas_de_dados: Math.max(0, sheet.getLastRow() - 1),
      como_abrir_relatorio: "acrescente ?id=ID_DO_RELATORIO ao endereço",
      outras_acoes: "?acao=lista (resumo) | ?acao=exportar (todos completos)"
    });
  }

  var lang = idiomaValido(p.lang);
  var rep = lerRelatorio(id);

  // O Resumo Executivo é calculado sobre o relatório ORIGINAL (em português),
  // porque a deteção de palavras-chave ("aprovado", "reprovado", ...) só faz
  // sentido nesse idioma. Onde o técnico tiver editado um quadrante à mão
  // (rep.resumoManual, vindo da app), essa versão substitui a sugestão
  // automática. As frases finais são traduzidas a seguir, junto com o resto
  // do relatório.
  if (rep) rep.resumoExecutivo = mesclarResumoExecutivo(gerarResumoExecutivo(rep), rep.resumoManual);

  if (rep && lang) rep = traduzirRelatorio(rep, lang);

  if (p.formato === "json") {
    return rep ? jsonOut({ ok: true, relatorio: rep, idioma: lang || "original" })
               : jsonOut({ ok: false, error: "Relatório '" + id + "' não encontrado." });
  }

  var html = rep ? paginaRelatorio(rep, lang) : paginaErro(id, lang);

  // Para descarregar como ficheiro a partir da app (ContentService permite CORS)
  if (p.formato === "ficheiro") {
    return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.HTML);
  }

  return HtmlService.createHtmlOutput(html)
    .setTitle(rep ? ("Relatório " + rep.id + " · MultiWasher") : "Relatório não encontrado")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Resumo de todos os relatórios, um por ID, mais recente primeiro
function listarRelatorios() {
  var sheet = getSheet();
  var ultima = sheet.getLastRow();
  if (ultima < 2) return [];

  var dados = sheet.getRange(2, 1, ultima - 1, HEADERS.length).getValues();
  var porId = {};
  var ordem = [];

  dados.forEach(function (r) {
    var id = String(r[1]).trim();
    if (!id) return;

    if (!porId[id]) {
      var nFotos = 0;
      try { nFotos = r[24] ? JSON.parse(r[24]).length : 0; } catch (err) { nFotos = 0; }
      porId[id] = {
        id: id,
        data: r[2] instanceof Date ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[2]),
        cliente: String(r[3]),
        setor: String(r[5]),
        tecnico: String(r[6]),
        maquina: String(r[7]),
        utensilios: Number(r[10]) || 1,
        fotos: nFotos
      };
      ordem.push(id);
    }
  });

  var lista = ordem.map(function (id) { return porId[id]; });
  lista.sort(function (a, b) { return String(b.data).localeCompare(String(a.data)); });
  return lista;
}

// Reagrupa as linhas de um ID num único objeto de relatório
function lerRelatorio(id) {
  var sheet = getSheet();
  var ultima = sheet.getLastRow();
  if (ultima < 2) return null;

  var dados = sheet.getRange(2, 1, ultima - 1, HEADERS.length).getValues();
  var linhas = dados.filter(function (r) {
    return String(r[1]).trim() === String(id).trim();
  });
  if (!linhas.length) return null;

  linhas.sort(function (a, b) { return (a[9] || 0) - (b[9] || 0); });
  var p = linhas[0];

  var fotos = [];
  try { fotos = p[24] ? JSON.parse(p[24]) : []; } catch (err) { fotos = []; }

  var resumoManual = {};
  try { resumoManual = p[27] ? JSON.parse(p[27]) : {}; } catch (err) { resumoManual = {}; }

  return {
    id: String(p[1]),
    data: p[2] instanceof Date ? Utilities.formatDate(p[2], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(p[2]),
    cliente: String(p[3]),
    clienteLogo: String(p[4]),
    setor: String(p[5]),
    tecnico: String(p[6]),
    maquina: String(p[7]),
    serialNumber: String(p[8]),
    utensilios: linhas.map(function (r) {
      return {
        tipo: String(r[11]), material: String(r[12]), sujidade: String(r[13]), objetivo: String(r[14]),
        vaporIni: String(r[15]), lavagem: String(r[16]), centrifInter: String(r[17]),
        enxague: String(r[18]), centrifFin: String(r[19]), vaporFin: String(r[20])
      };
    }),
    consideracoesPrevias: String(p[21]),
    consideracoes: String(p[22]),
    conclusao: String(p[23]),
    fotos: fotos,
    videoUrl: String(p[25]),
    pasta: String(p[26]),
    resumoManual: resumoManual,
    atualizado: p[0] instanceof Date ? Utilities.formatDate(p[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : ""
  };
}

/** ============================ RESUMO EXECUTIVO ============================ **/

/**
 * Gera automaticamente o "Resumo Executivo" (Pontos Fortes / Pontos a Melhorar /
 * Recomendações / Riscos & Alertas) a partir dos dados já existentes do ensaio —
 * sem inventar juízos técnicos sobre química ou materiais, só assinala o que
 * está ou não documentado e lê palavras-chave na conclusão/considerações.
 *
 * Espelha exatamente a função gerarResumoExecutivo() do index.html, para que
 * "Ver" (na app) e "Link" / "Descarregar" (servidos por este backend) mostrem
 * sempre o mesmo resumo.
 */
var RESUMO_PALAVRAS_POSITIVAS = ["aprovado", "sucesso", "100%", "eficaz", "eficiente", "sem danos", "excelente", "satisfatório", "satisfatoria", "conforme", "ótimo", "otimo", "muito bom", "bom desempenho"];
var RESUMO_PALAVRAS_NEGATIVAS = ["reprovado", "falha", "falhou", "não removeu", "nao removeu", "mancha", "resíduo", "residuo", "dano", "insuficiente", "problema", "dificuldade", "não conforme", "nao conforme", "incompleto", "ruído excessivo", "ruido excessivo", "corrosão", "corrosao"];

// Frases a remover antes de procurar as palavras-chave, para evitar falsos positivos
// por substring (ex: "sem danos" contém "dano"; "não conforme" contém "conforme").
var RESUMO_NEGACOES_POSITIVAS = ["não conforme", "nao conforme"];
var RESUMO_NEGACOES_NEGATIVAS = ["sem danos", "sem dano"];

var RESUMO_CAMPOS_PROGRAMA = ["vaporIni", "lavagem", "centrifInter", "enxague", "centrifFin", "vaporFin"];

function resumoContemAlguma(texto, lista, frasesARemover) {
  var t = texto;
  (frasesARemover || []).forEach(function (f) { t = t.split(f).join(" "); });
  return lista.some(function (k) { return t.indexOf(k) >= 0; });
}

function gerarResumoExecutivo(d) {
  var forcas = [], melhorar = [], recomendacoes = [], riscos = [];

  var conclusao = String(d.conclusao || "").toLowerCase();
  var consid = String(d.consideracoes || "").toLowerCase();
  var considPrev = String(d.consideracoesPrevias || "").toLowerCase();
  var utensilios = Array.isArray(d.utensilios) ? d.utensilios : [];
  var fotosValidas = (d.fotos || []).filter(function (f) { return f.url && String(f.url).trim() !== ""; });

  // Pontos fortes
  if (conclusao && resumoContemAlguma(conclusao, RESUMO_PALAVRAS_POSITIVAS, RESUMO_NEGACOES_POSITIVAS)) {
    forcas.push("Conclusão do ensaio regista um resultado positivo pelo técnico responsável.");
  }
  if (utensilios.length > 0) {
    var tiposUnicos = {}, materiaisUnicos = {}, nTipos = 0, nMateriais = 0;
    utensilios.forEach(function (u) {
      if (u.tipo && !tiposUnicos[u.tipo]) { tiposUnicos[u.tipo] = true; nTipos++; }
      if (u.material && !materiaisUnicos[u.material]) { materiaisUnicos[u.material] = true; nMateriais++; }
    });
    if (nTipos > 1 || nMateriais > 1) {
      forcas.push("Ensaio cobre " + nTipos + " tipo(s) de utensílio e " + nMateriais + " material(is) diferente(s), alargando a validação do processo.");
    }
    var comObjetivo = utensilios.filter(function (u) { return u.objetivo; }).length;
    if (comObjetivo === utensilios.length) {
      forcas.push("Todos os " + utensilios.length + " utensílio(s) testados têm objetivo de ensaio claramente definido.");
    }
  }
  if (fotosValidas.length > 0) {
    forcas.push("Ensaio documentado com " + fotosValidas.length + " fotografia(s) de evidência.");
  }
  if (forcas.length === 0) {
    forcas.push("Sem indicadores positivos suficientes nos dados atuais do relatório.");
  }

  // Pontos a melhorar
  if (conclusao && resumoContemAlguma(conclusao, RESUMO_PALAVRAS_NEGATIVAS, RESUMO_NEGACOES_NEGATIVAS)) {
    melhorar.push("Conclusão do ensaio menciona limitações ou resultados negativos a validar.");
  }
  if (consid && resumoContemAlguma(consid, RESUMO_PALAVRAS_NEGATIVAS, RESUMO_NEGACOES_NEGATIVAS)) {
    melhorar.push("Considerações durante/pós-ensaio referem problemas ou dificuldades observadas.");
  }
  var utensiliosIncompletos = utensilios.filter(function (u) {
    return !u.tipo || !u.material || !u.sujidade || !u.objetivo;
  }).length;
  if (utensiliosIncompletos > 0) {
    melhorar.push(utensiliosIncompletos + " utensílio(s) com campos por preencher (tipo, material, sujidade ou objetivo).");
  }
  var semPrograma = utensilios.filter(function (u) {
    return RESUMO_CAMPOS_PROGRAMA.every(function (k) { return !String(u[k] || "").trim(); });
  }).length;
  if (semPrograma > 0) {
    melhorar.push(semPrograma + " utensílio(s) sem programa de lavagem documentado.");
  }
  if (!considPrev) {
    melhorar.push("Não foram registadas considerações prévias ao ensaio.");
  }
  if (melhorar.length === 0) {
    melhorar.push("Nenhum ponto de melhoria identificado nos dados atuais do relatório.");
  }

  // Recomendações
  if (fotosValidas.length === 0) {
    recomendacoes.push("Adicionar fotografias de evidência para reforçar o relatório.");
  }
  if (!conclusao) {
    recomendacoes.push("Preencher a Conclusão e Avaliação Final antes de enviar o relatório.");
  }
  if (utensilios.length === 1) {
    recomendacoes.push("Considerar alargar o ensaio a mais utensílios/materiais para reforçar a validação do processo.");
  }
  if (!d.setor) {
    recomendacoes.push("Definir o setor/indústria do cliente para melhor contextualizar o relatório.");
  }
  if (recomendacoes.length === 0) {
    recomendacoes.push("Sem recomendações adicionais — relatório completo nos dados disponíveis.");
  }

  // Riscos & Alertas
  if (!conclusao) {
    riscos.push("Campo obrigatório 'Conclusão e Avaliação Final' está vazio — relatório incompleto.");
  }
  if (conclusao && resumoContemAlguma(conclusao, RESUMO_PALAVRAS_NEGATIVAS, RESUMO_NEGACOES_NEGATIVAS)) {
    riscos.push("Indicação de reprovação/falha na conclusão — validar antes de partilhar com o cliente.");
  }
  if (!d.cliente || !d.data) {
    riscos.push("Dados de identificação do ensaio incompletos (cliente ou data em falta).");
  }
  if (riscos.length === 0) {
    riscos.push("Sem riscos ou alertas identificados nos dados atuais do relatório.");
  }

  return { forcas: forcas, melhorar: melhorar, recomendacoes: recomendacoes, riscos: riscos };
}

// Substitui, quadrante a quadrante, a sugestão automática pelo que o técnico
// tiver editado à mão na app (rep.resumoManual). Um quadrante sem edição
// (chave ausente) mantém a sugestão automática.
function mesclarResumoExecutivo(auto, manual) {
  manual = manual || {};
  var final = {};
  ["forcas", "melhorar", "recomendacoes", "riscos"].forEach(function (campo) {
    final[campo] = Array.isArray(manual[campo]) ? manual[campo] : auto[campo];
  });
  return final;
}

/** ============================ DRIVE ============================ **/

function pastaDoRelatorio(id) {
  try {
    var raiz = DriveApp.getFolderById(PASTA_RAIZ_ID);
    var nome = "Relatorio_" + String(id).replace(/[^\w\-]/g, "_");
    var existentes = raiz.getFoldersByName(nome);

    if (existentes.hasNext()) {
      var pasta = existentes.next();
      // Reenvio do mesmo ID: limpa o conteúdo antigo
      var ficheiros = pasta.getFiles();
      while (ficheiros.hasNext()) ficheiros.next().setTrashed(true);
      return pasta;
    }

    var nova = raiz.createFolder(nome);
    definirPartilha(nova);
    return nova;
  } catch (err) {
    Logger.log("Não foi possível aceder à pasta do Drive: " + err);
    return null;
  }
}

// Manda a subpasta de fotos do relatório para o lixo do Drive
function mandarPastaParaLixo(id) {
  try {
    var raiz = DriveApp.getFolderById(PASTA_RAIZ_ID);
    var nome = "Relatorio_" + String(id).replace(/[^\w\-]/g, "_");
    var pastas = raiz.getFoldersByName(nome);
    if (!pastas.hasNext()) return false;
    pastas.next().setTrashed(true);
    return true;
  } catch (err) {
    Logger.log("Não foi possível mandar a pasta para o lixo: " + err);
    return false;
  }
}

/**
 * Guarda uma imagem base64 no Drive e devolve o URL.
 * Se já vier um URL normal (http...), devolve-o tal como está.
 */
function guardarImagem(pasta, nomeBase, valor) {
  if (!valor) return "";
  var s = String(valor);
  if (s.indexOf("data:") !== 0) return s;      // já é um link
  if (!pasta) return "";                        // sem pasta, não guarda base64 no Sheet

  try {
    var virgula = s.indexOf(",");
    var cabecalho = s.substring(0, virgula);
    var tipo = (cabecalho.match(/:(.*?);/) || [null, "image/jpeg"])[1];
    var ext = tipo.split("/")[1] || "jpg";
    if (ext === "jpeg") ext = "jpg";

    var bytes = Utilities.base64Decode(s.substring(virgula + 1));
    var blob = Utilities.newBlob(bytes, tipo, nomeBase + "." + ext);
    var ficheiro = pasta.createFile(blob);
    definirPartilha(ficheiro);

    return "https://drive.google.com/file/d/" + ficheiro.getId() + "/view";
  } catch (err) {
    Logger.log("Erro ao guardar imagem " + nomeBase + ": " + err);
    return "";
  }
}

function definirPartilha(item) {
  try {
    if (PARTILHA === "publico") {
      item.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } else if (PARTILHA === "dominio") {
      item.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    }
  } catch (err) {
    Logger.log("Não foi possível definir a partilha: " + err);
  }
}

// Extrai o ID do ficheiro de um endereço do Drive
function idDoDrive(url) {
  var m = String(url).match(/\/d\/([\w-]+)/) || String(url).match(/[?&]id=([\w-]+)/);
  return m ? m[1] : "";
}

// URL que mostra a imagem embutida numa página
function urlImagem(url) {
  var m = String(url).match(/\/d\/([\w-]+)/) || String(url).match(/[?&]id=([\w-]+)/);
  if (m) return "https://drive.google.com/thumbnail?id=" + m[1] + "&sz=w1600";
  return url;
}

/** ============================ PÁGINA HTML ============================ **/

function linkDoRelatorio(id, lang) {
  return ScriptApp.getService().getUrl() + "?id=" + encodeURIComponent(id) +
         (idiomaValido(lang) ? "&lang=" + lang : "");
}

function esc(v) {
  return String(v === null || v === undefined ? "" : v)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function ou(v) { return v && String(v).trim() !== "" ? esc(v) : "—"; }

// Lista de <li> do Resumo Executivo (já traduzidos, se aplicável)
function listaResumo(itens) {
  return (itens || []).map(function (t) { return "<li>" + esc(t) + "</li>"; }).join("");
}

// Bloco com os 4 quadrantes do Resumo Executivo
function resumoExecutivoHtml(r, L) {
  var re = r.resumoExecutivo || { forcas: [], melhorar: [], recomendacoes: [], riscos: [] };
  return ''
    + '<div class="mb-6">'
    +   '<h3 class="text-xs font-extrabold bg-slate-800 text-white px-3 py-1.5 rounded-t-lg tracking-wide uppercase mb-0">' + esc(L.resumo_titulo) + '</h3>'
    +   '<div class="grid grid-cols-1 sm:grid-cols-2 gap-3 border border-t-0 border-slate-200 rounded-b-lg p-3 bg-slate-50">'
    +     '<div class="bg-emerald-50 border border-emerald-200 rounded-lg p-3">'
    +       '<h4 class="text-[10px] font-black text-emerald-800 uppercase tracking-wide mb-1.5">' + esc(L.pontos_fortes) + '</h4>'
    +       '<ul class="text-[11px] text-emerald-900 space-y-1 list-disc list-inside">' + listaResumo(re.forcas) + '</ul>'
    +     '</div>'
    +     '<div class="bg-amber-50 border border-amber-200 rounded-lg p-3">'
    +       '<h4 class="text-[10px] font-black text-amber-800 uppercase tracking-wide mb-1.5">' + esc(L.pontos_melhorar) + '</h4>'
    +       '<ul class="text-[11px] text-amber-900 space-y-1 list-disc list-inside">' + listaResumo(re.melhorar) + '</ul>'
    +     '</div>'
    +     '<div class="bg-brand-50 border border-brand-200 rounded-lg p-3">'
    +       '<h4 class="text-[10px] font-black text-brand-800 uppercase tracking-wide mb-1.5">' + esc(L.recomendacoes) + '</h4>'
    +       '<ul class="text-[11px] text-brand-900 space-y-1 list-disc list-inside">' + listaResumo(re.recomendacoes) + '</ul>'
    +     '</div>'
    +     '<div class="bg-rose-50 border border-rose-200 rounded-lg p-3">'
    +       '<h4 class="text-[10px] font-black text-rose-800 uppercase tracking-wide mb-1.5">' + esc(L.riscos_alertas) + '</h4>'
    +       '<ul class="text-[11px] text-rose-900 space-y-1 list-disc list-inside">' + listaResumo(re.riscos) + '</ul>'
    +     '</div>'
    +   '</div>'
    + '</div>';
}

function paginaRelatorio(r, lang) {
  var L = etiquetas(lang);
  var logoCliente = r.clienteLogo
    ? '<img src="' + esc(urlImagem(r.clienteLogo)) + '" alt="' + esc(L.cliente) + '" class="h-12 w-auto object-contain">'
    : '';

  var utensilios = r.utensilios.map(function (u, i) {
    return ''
      + '<div class="mb-6">'
      +   '<h3 class="text-xs font-extrabold bg-brand-800 text-white px-3 py-1.5 rounded-t-lg tracking-wide uppercase">'
      +     esc(L.utensilio_de.replace("{n}", i + 1).replace("{t}", r.utensilios.length))
      +   '</h3>'
      +   '<div class="border border-slate-200 border-t-0 rounded-b-lg p-3 space-y-2 text-xs">'
      +     '<div class="grid grid-cols-1 sm:grid-cols-4 gap-2 bg-slate-50 p-2 rounded">'
      +       '<div><strong class="text-slate-600">' + esc(L.utensilio) + ':</strong> ' + ou(u.tipo) + '</div>'
      +       '<div><strong class="text-slate-600">' + esc(L.material) + ':</strong> ' + ou(u.material) + '</div>'
      +       '<div><strong class="text-slate-600">' + esc(L.sujidade) + ':</strong> ' + ou(u.sujidade) + '</div>'
      +       '<div><strong class="text-slate-600">' + esc(L.objetivo) + ':</strong> ' + ou(u.objetivo) + '</div>'
      +     '</div>'
      +     '<div class="grid grid-cols-2 sm:grid-cols-6 gap-2 text-[10px] text-center pt-1">'
      +       celula(L.vapor_ini, u.vaporIni) + celula(L.lavagem, u.lavagem)
      +       celula(L.centr_inter, u.centrifInter) + celula(L.enxaguamento, u.enxague)
      +       celula(L.centr_final, u.centrifFin) + celula(L.vapor_final, u.vaporFin)
      +     '</div>'
      +   '</div>'
      + '</div>';
  }).join("");

  var comFoto = (r.fotos || []).filter(function (f) { return f.url || f.desc; });
  var fotos = comFoto.length
    ? comFoto.map(function (f) {
        var img = f.url
          ? '<a href="' + esc(f.url) + '" target="_blank" rel="noopener">'
          +   '<img src="' + esc(urlImagem(f.url)) + '" data-drive="' + esc(idDoDrive(f.url)) + '"'
          +   ' class="foto-evidencia w-full h-28 object-cover hover:opacity-90 transition"'
          +   ' alt="' + f.n + '" referrerpolicy="no-referrer">'
          + '</a>'
          : '<div class="w-full h-28 bg-slate-200 flex items-center justify-center text-slate-400 text-xl">—</div>';
        return '<div class="border border-slate-200 rounded-lg overflow-hidden bg-slate-50">'
          + img
          + '<div class="p-1.5 text-[10px] text-slate-700 font-medium leading-tight text-center">' + ou(f.desc) + '</div>'
          + '</div>';
      }).join("")
    : '<div class="col-span-full text-center text-slate-400 py-4 text-xs italic">' + esc(L.sem_fotos) + '</div>';

  var video = r.videoUrl
    ? '<p class="text-xs mt-3"><strong class="text-slate-600">' + esc(L.video) + ':</strong> '
      + '<a href="' + esc(r.videoUrl) + '" target="_blank" rel="noopener" class="text-brand-600 underline break-all">' + esc(r.videoUrl) + '</a></p>'
    : '';

  return ''
+ '<!DOCTYPE html><html lang="' + (idiomaValido(lang) || "pt") + '"><head><meta charset="UTF-8">'
+ '<script src="https://cdn.tailwindcss.com"></script>'
+ '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">'
+ '<script>tailwind.config={theme:{extend:{fontFamily:{sans:["Inter","sans-serif"]},colors:{brand:{50:"#f0f7ff",100:"#e0effe",200:"#bae0fd",300:"#7cc8fc",400:"#36aff8",500:"#0c94eb",600:"#0076c9",700:"#025ea7",800:"#06508a",900:"#0b4372"}}}}}</script>'
+ '<style>@media print{body{background:#fff!important}.no-print{display:none!important}.print-container{box-shadow:none!important;border:none!important;padding:0!important}}</style>'
+ '</head><body class="bg-slate-100 font-sans text-slate-800 antialiased pb-10">'

+ '<header class="bg-slate-900 text-white shadow-md no-print">'
+   '<div class="max-w-5xl mx-auto px-4 py-3 flex flex-col sm:flex-row items-center justify-between gap-3">'
+     '<div><p class="text-xs font-black uppercase tracking-wider">' + esc(L.titulo) + '</p>'
+     '<p class="text-[11px] text-slate-300">' + esc(r.id) + ' · ' + ou(r.cliente) + ' · ' + esc(L.atualizado) + ' ' + esc(r.atualizado) + '</p></div>'
+     '<button id="btn-imprimir" onclick="imprimirComFotos()" class="px-4 py-1.5 bg-brand-600 hover:bg-brand-500 text-white text-xs font-bold rounded-lg shadow transition">' + esc(L.imprimir) + '</button>'
+   '</div>'
+   barraIdiomas(r.id, lang, L)
+ '</header>'

+ '<div class="max-w-5xl mx-auto px-3 sm:px-4 pt-5">'
+ '<div class="bg-white border border-slate-300 rounded-2xl shadow-xl p-6 sm:p-10 print-container">'

+   '<div class="flex items-center justify-between border-b-2 border-brand-800 pb-4 mb-6">'
+     '<img src="' + LOGO_SOMENGIL + '" alt="Somengil" class="h-12 w-auto object-contain">'
+     '<div class="text-right">'
+       '<h1 class="text-lg font-black text-brand-800 uppercase tracking-wide">' + esc(L.titulo) + '</h1>'
+       '<p class="text-xs text-slate-500">' + esc(r.id) + '</p>'
+     '</div>'
+     (logoCliente ? '<div class="pl-4">' + logoCliente + '</div>' : '')
+   '</div>'

+   '<div class="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs mb-6">'
+     campo(L.cliente, r.cliente) + campo(L.setor, r.setor)
+     campo(L.data, r.data) + campo(L.tecnico, r.tecnico)
+     campo(L.equipamento, r.maquina) + campo(L.serie, r.serialNumber)
+     campo(L.n_utensilios, String(r.utensilios.length))
+   '</div>'

+   resumoExecutivoHtml(r, L)

+   utensilios

+   '<div class="mb-6 space-y-3 text-xs">'
+     bloco(L.cons_previas, r.consideracoesPrevias)
+     bloco(L.cons_praticas, r.consideracoes)
+     '<div class="bg-brand-50 border border-brand-200 p-3 rounded-lg">'
+       '<p class="font-black text-brand-800 uppercase text-[11px] tracking-wide mb-1">' + esc(L.conclusao) + '</p>'
+       '<p class="text-slate-700 leading-relaxed">' + ou(r.conclusao) + '</p>'
+     '</div>'
+   '</div>'

+   '<h3 class="text-xs font-extrabold bg-brand-800 text-white px-3 py-1.5 rounded-t-lg tracking-wide uppercase">' + esc(L.evidencias) + '</h3>'
+   '<div class="border border-slate-200 border-t-0 rounded-b-lg p-3">'
+     '<div class="grid grid-cols-2 sm:grid-cols-5 gap-3">' + fotos + '</div>'
+     video
+   '</div>'

+   '<p class="text-[10px] text-slate-400 text-center mt-8 pt-4 border-t border-slate-200">'
+     esc(L.rodape.replace("{id}", r.id))
+     (idiomaValido(lang) ? '<br>' + esc(L.aviso_traducao) : '')
+   '</p>'

+ '</div></div>'

+ '<script>'
+ '(function(){'
+   'var tentativas={};'
+   'document.querySelectorAll("img.foto-evidencia").forEach(function(img){'
+     'img.addEventListener("error",function(){'
+       'var id=img.getAttribute("data-drive"); if(!id) return;'
+       'var n=(tentativas[id]||0)+1; tentativas[id]=n;'
+       'if(n===1){img.src="https://lh3.googleusercontent.com/d/"+id+"=w1600";}'
+       'else if(n===2){img.src="https://drive.google.com/thumbnail?id="+id+"&sz=w800";}'
+     '});'
+   '});'
+   'window.imprimirComFotos=function(){'
+     'var btn=document.getElementById("btn-imprimir");'
+     'var imgs=[].slice.call(document.querySelectorAll("img"));'
+     'var faltam=imgs.filter(function(i){return !i.complete;});'
+     'if(!faltam.length){window.print();return;}'
+     'if(btn){btn.disabled=true;btn.textContent="..." ;}'
+     'var pendentes=faltam.length, feito=false;'
+     'function pronto(){if(feito)return;feito=true;if(btn){btn.disabled=false;btn.textContent=' + JSON.stringify(L.imprimir) + ';}window.print();}'
+     'faltam.forEach(function(i){'
+       'function conta(){if(--pendentes<=0)pronto();}'
+       'i.addEventListener("load",conta);i.addEventListener("error",conta);'
+     '});'
+     'setTimeout(pronto,8000);'   // nunca deixa o utilizador à espera
+   '};'
+ '})();'
+ '<\/script>'

+ '</body></html>';
}

// Barra para o destinatário mudar de idioma na própria página
function barraIdiomas(id, lang, L) {
  var atual = idiomaValido(lang) || "";
  var base = ScriptApp.getService().getUrl() + "?id=" + encodeURIComponent(id);

  var links = Object.keys(IDIOMAS).map(function (k) {
    var activo = (k === atual);
    return '<a href="' + base + '&lang=' + k + '" class="px-2.5 py-1 rounded-md text-[11px] font-bold transition '
      + (activo ? 'bg-brand-600 text-white' : 'text-slate-300 hover:bg-white/10') + '">'
      + k.toUpperCase() + '</a>';
  }).join("");

  var original = '<a href="' + base + '" class="px-2.5 py-1 rounded-md text-[11px] font-bold transition '
    + (atual ? 'text-slate-300 hover:bg-white/10' : 'bg-slate-600 text-white') + '">Original</a>';

  return '<div class="border-t border-white/10">'
    + '<div class="max-w-5xl mx-auto px-4 py-2 flex flex-wrap items-center gap-1.5">'
    + '<span class="text-[10px] uppercase tracking-wider text-slate-400 font-bold mr-1">' + esc(L.idioma) + '</span>'
    + original + links
    + '</div></div>';
}

function campo(rotulo, valor) {
  return '<div class="bg-slate-50 border border-slate-200 rounded p-2">'
    + '<span class="block text-[10px] font-bold text-slate-500 uppercase">' + esc(rotulo) + '</span>'
    + '<span class="font-semibold text-slate-800">' + ou(valor) + '</span></div>';
}

function celula(rotulo, valor) {
  return '<div class="bg-slate-100 p-1.5 rounded border border-slate-200">'
    + '<span class="block font-bold text-slate-500">' + esc(rotulo) + '</span>'
    + '<span class="font-semibold">' + ou(valor) + '</span></div>';
}

function bloco(titulo, texto) {
  return '<div class="bg-slate-50 border border-slate-200 p-3 rounded-lg">'
    + '<p class="font-black text-slate-600 uppercase text-[11px] tracking-wide mb-1">' + esc(titulo) + '</p>'
    + '<p class="text-slate-700 leading-relaxed">' + ou(texto) + '</p></div>';
}

function paginaErro(id, lang) {
  var L = etiquetas(lang);
  return '<!DOCTYPE html><html lang="' + (idiomaValido(lang) || "pt") + '"><head><meta charset="UTF-8">'
    + '<script src="https://cdn.tailwindcss.com"></script></head>'
    + '<body class="bg-slate-100 min-h-screen flex items-center justify-center p-6">'
    + '<div class="bg-white border border-slate-200 rounded-2xl shadow p-8 max-w-md text-center">'
    + '<p class="text-4xl mb-3">🔍</p>'
    + '<h1 class="text-base font-black text-slate-800 mb-2">' + esc(L.nao_encontrado) + '</h1>'
    + '<p class="text-sm text-slate-600">' + esc(L.nao_existe.replace("{id}", id)) + '</p>'
    + '</div></body></html>';
}

/** ============================ AUXILIARES ============================ **/

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normalizarUtensilios(d) {
  if (d.utensilios && d.utensilios.length) return d.utensilios;

  var lista = [];
  [1, 2].forEach(function (n) {
    var u = {
      tipo: d["utensilio" + n] || "", material: d["material" + n] || "",
      sujidade: d["sujidade" + n] || "", objetivo: d["objetivo" + n] || "",
      vaporIni: d["prog" + n + "VaporIni"] || "", lavagem: d["prog" + n + "Lavagem"] || "",
      centrifInter: d["prog" + n + "CentrifInter"] || "", enxague: d["prog" + n + "Enxague"] || "",
      centrifFin: d["prog" + n + "CentrifFin"] || "", vaporFin: d["prog" + n + "VaporFin"] || ""
    };
    var preenchido = Object.keys(u).some(function (k) { return String(u[k]).trim() !== ""; });
    if (preenchido) lista.push(u);
  });

  return lista.length ? lista : [{}];
}

function apagarLinhasDoId(sheet, id) {
  if (!id) return 0;
  var ultima = sheet.getLastRow();
  if (ultima < 2) return 0;

  var ids = sheet.getRange(2, COL_ID, ultima - 1, 1).getValues();
  var apagadas = 0;
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 2);
      apagadas++;
    }
  }
  return apagadas;
}

function txt(v) {
  if (v === null || v === undefined) return "";
  var s = String(v);
  return s.length > MAX_CELL ? s.substring(0, MAX_CELL) + " …[cortado]" : s;
}

function pad2(n) { return n < 10 ? "0" + n : String(n); }

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


var IDIOMAS = {"pt": "Português", "en": "English", "es": "Español", "fr": "Français", "de": "Deutsch"};
var LABELS = {
  "pt": {
    "titulo": "Relatório de Ensaio de Lavagem",
    "imprimir": "Imprimir / Guardar PDF",
    "atualizado": "atualizado a",
    "cliente": "Cliente",
    "setor": "Setor / Indústria",
    "data": "Data do Ensaio",
    "tecnico": "Técnico Responsável",
    "equipamento": "Equipamento",
    "serie": "Nº de Série",
    "n_utensilios": "Utensílios ensaiados",
    "utensilio_de": "Utensílio {n} de {t}: Especificação do Ensaio",
    "utensilio": "Utensílio",
    "material": "Material",
    "sujidade": "Sujidade",
    "objetivo": "Objetivo",
    "vapor_ini": "Vapor Inicial",
    "lavagem": "Lavagem",
    "centr_inter": "Centr. Interm.",
    "enxaguamento": "Enxaguamento",
    "centr_final": "Centr. Final",
    "vapor_final": "Vapor Final",
    "cons_previas": "Considerações Prévias",
    "cons_praticas": "Considerações Práticas",
    "conclusao": "Conclusão Final",
    "evidencias": "Evidências Fotográficas",
    "sem_fotos": "Nenhuma fotografia de evidência anexada.",
    "video": "Vídeo do ensaio",
    "resumo_titulo": "Resumo Executivo do Ensaio",
    "pontos_fortes": "Pontos Fortes",
    "pontos_melhorar": "Pontos a Melhorar",
    "recomendacoes": "Recomendações",
    "riscos_alertas": "Riscos & Alertas",
    "rodape": "Somengil · MultiWasher — documento gerado automaticamente a partir do registo do ensaio {id}.",
    "idioma": "Idioma",
    "aviso_traducao": "Tradução automática dos textos escritos pelo técnico.",
    "nao_encontrado": "Relatório não encontrado",
    "nao_existe": "Não existe nenhum registo com o ID {id}. Confirme o endereço."
  },
  "en": {
    "titulo": "Washing Trial Report",
    "imprimir": "Print / Save as PDF",
    "atualizado": "updated on",
    "cliente": "Customer",
    "setor": "Sector / Industry",
    "data": "Trial date",
    "tecnico": "Technician",
    "equipamento": "Equipment",
    "serie": "Serial number",
    "n_utensilios": "Utensils tested",
    "utensilio_de": "Utensil {n} of {t}: Trial specification",
    "utensilio": "Utensil",
    "material": "Material",
    "sujidade": "Soiling",
    "objetivo": "Objective",
    "vapor_ini": "Initial steam",
    "lavagem": "Wash",
    "centr_inter": "Interm. spin",
    "enxaguamento": "Rinse",
    "centr_final": "Final spin",
    "vapor_final": "Final steam",
    "cons_previas": "Preliminary notes",
    "cons_praticas": "Practical notes",
    "conclusao": "Conclusion",
    "evidencias": "Photographic evidence",
    "sem_fotos": "No evidence photographs attached.",
    "video": "Trial video",
    "resumo_titulo": "Executive Summary",
    "pontos_fortes": "Strengths",
    "pontos_melhorar": "Areas to Improve",
    "recomendacoes": "Recommendations",
    "riscos_alertas": "Risks & Alerts",
    "rodape": "Somengil · MultiWasher — document generated automatically from trial record {id}.",
    "idioma": "Language",
    "aviso_traducao": "The technician's notes were machine-translated.",
    "nao_encontrado": "Report not found",
    "nao_existe": "There is no record with ID {id}. Please check the address."
  },
  "es": {
    "titulo": "Informe de Ensayo de Lavado",
    "imprimir": "Imprimir / Guardar PDF",
    "atualizado": "actualizado el",
    "cliente": "Cliente",
    "setor": "Sector / Industria",
    "data": "Fecha del ensayo",
    "tecnico": "Técnico responsable",
    "equipamento": "Equipo",
    "serie": "Nº de serie",
    "n_utensilios": "Utensilios ensayados",
    "utensilio_de": "Utensilio {n} de {t}: Especificación del ensayo",
    "utensilio": "Utensilio",
    "material": "Material",
    "sujidade": "Suciedad",
    "objetivo": "Objetivo",
    "vapor_ini": "Vapor inicial",
    "lavagem": "Lavado",
    "centr_inter": "Centrif. interm.",
    "enxaguamento": "Aclarado",
    "centr_final": "Centrif. final",
    "vapor_final": "Vapor final",
    "cons_previas": "Consideraciones previas",
    "cons_praticas": "Consideraciones prácticas",
    "conclusao": "Conclusión final",
    "evidencias": "Evidencias fotográficas",
    "sem_fotos": "No se adjuntaron fotografías de evidencia.",
    "video": "Vídeo del ensayo",
    "resumo_titulo": "Resumen Ejecutivo del Ensayo",
    "pontos_fortes": "Puntos Fuertes",
    "pontos_melhorar": "Puntos a Mejorar",
    "recomendacoes": "Recomendaciones",
    "riscos_alertas": "Riesgos y Alertas",
    "rodape": "Somengil · MultiWasher — documento generado automáticamente a partir del registro del ensayo {id}.",
    "idioma": "Idioma",
    "aviso_traducao": "Los textos del técnico se han traducido automáticamente.",
    "nao_encontrado": "Informe no encontrado",
    "nao_existe": "No existe ningún registro con el ID {id}. Compruebe la dirección."
  },
  "fr": {
    "titulo": "Rapport d'Essai de Lavage",
    "imprimir": "Imprimer / Enregistrer en PDF",
    "atualizado": "mis à jour le",
    "cliente": "Client",
    "setor": "Secteur / Industrie",
    "data": "Date de l'essai",
    "tecnico": "Technicien responsable",
    "equipamento": "Équipement",
    "serie": "N° de série",
    "n_utensilios": "Ustensiles testés",
    "utensilio_de": "Ustensile {n} sur {t} : spécification de l'essai",
    "utensilio": "Ustensile",
    "material": "Matériau",
    "sujidade": "Salissure",
    "objetivo": "Objectif",
    "vapor_ini": "Vapeur initiale",
    "lavagem": "Lavage",
    "centr_inter": "Essorage interm.",
    "enxaguamento": "Rinçage",
    "centr_final": "Essorage final",
    "vapor_final": "Vapeur finale",
    "cons_previas": "Observations préalables",
    "cons_praticas": "Observations pratiques",
    "conclusao": "Conclusion finale",
    "evidencias": "Preuves photographiques",
    "sem_fotos": "Aucune photographie jointe.",
    "video": "Vidéo de l'essai",
    "resumo_titulo": "Résumé Exécutif de l'Essai",
    "pontos_fortes": "Points Forts",
    "pontos_melhorar": "Points à Améliorer",
    "recomendacoes": "Recommandations",
    "riscos_alertas": "Risques et Alertes",
    "rodape": "Somengil · MultiWasher — document généré automatiquement à partir de l'enregistrement d'essai {id}.",
    "idioma": "Langue",
    "aviso_traducao": "Les textes du technicien ont été traduits automatiquement.",
    "nao_encontrado": "Rapport introuvable",
    "nao_existe": "Aucun enregistrement ne correspond à l'identifiant {id}. Vérifiez l'adresse."
  },
  "de": {
    "titulo": "Waschversuchsbericht",
    "imprimir": "Drucken / Als PDF speichern",
    "atualizado": "aktualisiert am",
    "cliente": "Kunde",
    "setor": "Branche",
    "data": "Datum des Versuchs",
    "tecnico": "Zuständiger Techniker",
    "equipamento": "Anlage",
    "serie": "Seriennummer",
    "n_utensilios": "Geprüfte Utensilien",
    "utensilio_de": "Utensil {n} von {t}: Versuchsspezifikation",
    "utensilio": "Utensil",
    "material": "Material",
    "sujidade": "Verschmutzung",
    "objetivo": "Ziel",
    "vapor_ini": "Dampf am Anfang",
    "lavagem": "Waschen",
    "centr_inter": "Zwischenschleudern",
    "enxaguamento": "Spülen",
    "centr_final": "Endschleudern",
    "vapor_final": "Dampf am Ende",
    "cons_previas": "Vorbemerkungen",
    "cons_praticas": "Praktische Anmerkungen",
    "conclusao": "Schlussfolgerung",
    "evidencias": "Fotodokumentation",
    "sem_fotos": "Keine Nachweisfotos vorhanden.",
    "video": "Video des Versuchs",
    "resumo_titulo": "Zusammenfassung des Versuchs",
    "pontos_fortes": "Stärken",
    "pontos_melhorar": "Verbesserungspotenzial",
    "recomendacoes": "Empfehlungen",
    "riscos_alertas": "Risiken & Hinweise",
    "rodape": "Somengil · MultiWasher — automatisch erzeugtes Dokument aus dem Versuchsdatensatz {id}.",
    "idioma": "Sprache",
    "aviso_traducao": "Die Texte des Technikers wurden maschinell übersetzt.",
    "nao_encontrado": "Bericht nicht gefunden",
    "nao_existe": "Es existiert kein Datensatz mit der ID {id}. Bitte prüfen Sie die Adresse."
  },
    "it": {
    "titulo": "Rapporto di Prova di Lavaggio",
    "imprimir": "Stampa / Salva come PDF",
    "atualizado": "aggiornato il",
    "cliente": "Cliente",
    "setor": "Settore / Industria",
    "data": "Data della prova",
    "tecnico": "Tecnico responsabile",
    "equipamento": "Attrezzatura",
    "serie": "Numero di serie",
    "n_utensilios": "Utensili testati",
    "utensilio_de": "Utensile {n} di {t}: Specifica della prova",
    "utensilio": "Utensile",
    "material": "Materiale",
    "sujidade": "Sporco",
    "objetivo": "Obiettivo",
    "vapor_ini": "Vapore iniziale",
    "lavagem": "Lavaggio",
    "centr_inter": "Centrifuga interm.",
    "enxaguamento": "Risciacquo",
    "centr_final": "Centrifuga finale",
    "vapor_final": "Vapore finale",
    "cons_previas": "Considerazioni preliminari",
    "cons_praticas": "Considerazioni pratiche",
    "conclusao": "Conclusione finale",
    "evidencias": "Evidenze fotografiche",
    "sem_fotos": "Nessuna fotografia di evidenza allegata.",
    "video": "Video della prova",
    "resumo_titulo": "Riepilogo Esecutivo della Prova",
    "pontos_fortes": "Punti di Forza",
    "pontos_melhorar": "Punti da Migliorare",
    "recomendacoes": "Raccomandazioni",
    "riscos_alertas": "Rischi e Avvisi",
    "rodape": "Somengil · MultiWasher — documento generato automaticamente dal registro della prova {id}.",
    "idioma": "Lingua",
    "aviso_traducao": "I testi del tecnico sono stati tradotti automaticamente.",
    "nao_encontrado": "Rapporto non trovato",
    "nao_existe": "Non esiste alcun registro con l'ID {id}. Verificare l'indirizzo."
  }
};

/** ============ TRADUÇÃO ============ **/

/**
 * Traduz os textos escritos pelo técnico com o tradutor do Apps Script.
 * As etiquetas fixas vêm do dicionário LABELS, não do tradutor.
 *
 * Cada relatório é traduzido uma única vez por idioma: o resultado fica
 * guardado na folha _Traducoes, para não gastar quota a cada visita.
 */
var FOLHA_TRADUCOES = "_Traducoes";

function idiomaValido(lang) {
  return lang && LABELS[lang] ? lang : null;
}

function etiquetas(lang) {
  return LABELS[idiomaValido(lang) || "pt"];
}

function traduzir(texto, lang) {
  var s = String(texto || "");
  if (!s.trim()) return s;
  try {
    return LanguageApp.translate(s, "", lang);
  } catch (err) {
    Logger.log("Tradução falhou (quota?): " + err);
    return s;   // fica o original, melhor do que ficar vazio
  }
}

function traduzirRelatorio(rep, lang) {
  if (!idiomaValido(lang)) return rep;

  var chave = rep.id + "|" + lang;
  var guardado = lerTraducao(chave);
  if (guardado) return guardado;

  var t = JSON.parse(JSON.stringify(rep));

  // não se traduzem nomes próprios, datas, números de série nem endereços
  t.setor = traduzir(t.setor, lang);
  t.consideracoesPrevias = traduzir(t.consideracoesPrevias, lang);
  t.consideracoes = traduzir(t.consideracoes, lang);
  t.conclusao = traduzir(t.conclusao, lang);

  (t.utensilios || []).forEach(function (u) {
    ["tipo", "material", "sujidade", "objetivo", "vaporIni", "lavagem",
     "centrifInter", "enxague", "centrifFin", "vaporFin"].forEach(function (k) {
      u[k] = traduzir(u[k], lang);
    });
  });

  (t.fotos || []).forEach(function (f) { f.desc = traduzir(f.desc, lang); });

  // Resumo Executivo: as frases já foram geradas em português por
  // gerarResumoExecutivo(); aqui só as traduzimos, como o resto do texto livre.
  if (t.resumoExecutivo) {
    ["forcas", "melhorar", "recomendacoes", "riscos"].forEach(function (k) {
      t.resumoExecutivo[k] = (t.resumoExecutivo[k] || []).map(function (frase) {
        return traduzir(frase, lang);
      });
    });
  }

  t.idioma = lang;
  gravarTraducao(chave, t);
  return t;
}

function folhaTraducoes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(FOLHA_TRADUCOES);
  if (!f) {
    f = ss.insertSheet(FOLHA_TRADUCOES);
    f.appendRow(["Chave (id|idioma)", "JSON traduzido", "Quando"]);
    f.setFrozenRows(1);
    f.hideSheet();
  }
  return f;
}

function lerTraducao(chave) {
  try {
    var cache = CacheService.getScriptCache().get(chave);
    if (cache) return JSON.parse(cache);

    var f = folhaTraducoes();
    var ultima = f.getLastRow();
    if (ultima < 2) return null;

    var chaves = f.getRange(2, 1, ultima - 1, 2).getValues();
    for (var i = 0; i < chaves.length; i++) {
      if (String(chaves[i][0]) === chave) {
        CacheService.getScriptCache().put(chave, chaves[i][1], 21600);
        return JSON.parse(chaves[i][1]);
      }
    }
  } catch (err) {
    Logger.log("Cache de tradução ilegível: " + err);
  }
  return null;
}

function gravarTraducao(chave, obj) {
  try {
    var json = JSON.stringify(obj);
    if (json.length > MAX_CELL) return;          // demasiado grande para a célula
    folhaTraducoes().appendRow([chave, json, new Date()]);
    CacheService.getScriptCache().put(chave, json, 21600);
  } catch (err) {
    Logger.log("Não foi possível guardar a tradução: " + err);
  }
}

/** Apaga as traduções guardadas (para voltar a traduzir do zero). */
function limparTraducoes() {
  var f = folhaTraducoes();
  if (f.getLastRow() > 1) f.getRange(2, 1, f.getLastRow() - 1, 3).clearContent();
  CacheService.getScriptCache().removeAll([]);
  Logger.log("Traduções guardadas apagadas.");
}

/** ============ RECUPERAR AS FOTOS DO APPSHEET ANTIGO ============ **/

/**
 * Os relatórios importados têm links do tipo
 *   https://www.appsheet.com/template/gettablefileurl?...&fileName=INTRO_Images/6.FOTO%201.jpg
 * Esses endereços exigem sessão do AppSheet: num <img> aparecem sempre quebrados.
 *
 * Os ficheiros originais estão numa pasta do Drive (normalmente "INTRO_Images",
 * ao lado da folha de cálculo da app antiga). Esta função procura cada ficheiro
 * por nome nessa pasta e reescreve o JSON de fotos com o link do Drive.
 *
 * COMO USAR
 *  1. PASTA_APPSHEET_ID já está preenchido com a pasta INTRO_Images.
 *  2. Correr diagnosticoFotos() para ver o ponto de partida.
 *  3. Correr recuperarFotosDoAppSheet(). Partilha a pasta uma vez e reescreve o
 *     JSON linha a linha. Se der timeout, guarda onde ficou: basta voltar a
 *     correr até o registo dizer "CONCLUÍDO".
 *  4. Correr diagnosticoFotos() outra vez para confirmar.
 *  5. Para começar de novo do início: reiniciarRecuperacaoFotos().
 */

var PASTA_APPSHEET_ID = "1d4WNXJTe8MtgF_yaM2qM1zu1J9J8ecwM"; // INTRO_Images (app antiga)
var LOTE_LINHAS = 500;  // linhas por execução; se der timeout, baixar para 60 e voltar a correr

function recuperarFotosDoAppSheet() {
  if (PASTA_APPSHEET_ID.indexOf("COLE_AQUI") === 0) {
    throw new Error("Falta preencher PASTA_APPSHEET_ID com o ID da pasta das fotos antigas.");
  }

  // Uma única partilha na pasta: os ficheiros herdam-na. Muito mais rápido
  // do que partilhar 800 ficheiros um a um.
  try {
    definirPartilha(DriveApp.getFolderById(PASTA_APPSHEET_ID));
    Logger.log("Partilha aplicada à pasta INTRO_Images (modo: " + PARTILHA + ").");
  } catch (err) {
    Logger.log("AVISO: não foi possível alterar a partilha da pasta (" + err + "). " +
               "Se as imagens não abrirem a terceiros, partilhe a pasta à mão.");
  }

  var props = PropertiesService.getScriptProperties();
  var inicio = Number(props.getProperty("recup_linha") || 2);
  var sheet = getSheet();
  var ultima = sheet.getLastRow();
  if (ultima < 2) { Logger.log("Folha vazia."); return; }

  var mapa = mapaDeFicheiros(PASTA_APPSHEET_ID);
  Logger.log("Ficheiros encontrados na pasta: " + Object.keys(mapa).length);

  var fim = Math.min(ultima, inicio + LOTE_LINHAS - 1);
  var intervalo = sheet.getRange(inicio, 25, fim - inicio + 1, 1);   // coluna Y = Fotos (JSON)
  var valores = intervalo.getValues();

  var recuperadas = 0, naoEncontradas = 0, jaOk = 0;

  for (var i = 0; i < valores.length; i++) {
    var bruto = valores[i][0];
    if (!bruto) continue;

    var fotos;
    try { fotos = JSON.parse(bruto); } catch (err) { continue; }
    var mudou = false;

    fotos.forEach(function (f) {
      var url = String(f.url || "");
      if (!url) return;
      if (url.indexOf("drive.google.com") >= 0) { jaOk++; return; }
      if (url.indexOf("gettablefileurl") < 0) return;

      var nome = nomeDoFicheiro(url);
      var ficheiro = nome ? mapa[nome.toLowerCase()] : null;

      if (ficheiro) {
        f.url = "https://drive.google.com/file/d/" + ficheiro.getId() + "/view";
        recuperadas++;
        mudou = true;
      } else {
        naoEncontradas++;
      }
    });

    if (mudou) valores[i][0] = JSON.stringify(fotos);
  }

  intervalo.setValues(valores);

  if (fim >= ultima) {
    props.deleteProperty("recup_linha");
    Logger.log("CONCLUÍDO. Última execução: linhas " + inicio + " a " + fim +
               " | recuperadas: " + recuperadas + " | sem ficheiro: " + naoEncontradas + " | já em Drive: " + jaOk);
  } else {
    props.setProperty("recup_linha", String(fim + 1));
    Logger.log("Linhas " + inicio + " a " + fim + " tratadas | recuperadas: " + recuperadas +
               " | sem ficheiro: " + naoEncontradas + " | já em Drive: " + jaOk +
               "\nFALTA CONTINUAR: voltar a correr recuperarFotosDoAppSheet().");
  }
}

function reiniciarRecuperacaoFotos() {
  PropertiesService.getScriptProperties().deleteProperty("recup_linha");
  Logger.log("Contador reiniciado: a próxima execução começa na linha 2.");
}

// Índice nome -> ficheiro, incluindo subpastas
function mapaDeFicheiros(pastaId) {
  var mapa = {};
  var pastas = [DriveApp.getFolderById(pastaId)];

  while (pastas.length) {
    var p = pastas.pop();
    var fs = p.getFiles();
    while (fs.hasNext()) {
      var f = fs.next();
      mapa[f.getName().toLowerCase()] = f;
    }
    var subs = p.getFolders();
    while (subs.hasNext()) pastas.push(subs.next());
  }
  return mapa;
}

// Extrai "6.FOTO 1.154919.jpg" do parâmetro fileName do link do AppSheet
function nomeDoFicheiro(url) {
  var m = String(url).match(/[?&]fileName=([^&]+)/);
  if (!m) return "";
  var caminho = decodeURIComponent(m[1]);
  return caminho.split("/").pop();
}

/**
 * Alternativa, se os ficheiros originais já não existirem: limpa os endereços
 * quebrados e deixa só as descrições, para o relatório não mostrar imagens
 * partidas. Também remove o placeholder da app antiga (static.wixstatic.com).
 */
function limparFotosQuebradas() {
  var sheet = getSheet();
  var ultima = sheet.getLastRow();
  if (ultima < 2) return;

  var intervalo = sheet.getRange(2, 25, ultima - 1, 1);
  var valores = intervalo.getValues();
  var limpas = 0, removidas = 0;

  for (var i = 0; i < valores.length; i++) {
    if (!valores[i][0]) continue;
    var fotos;
    try { fotos = JSON.parse(valores[i][0]); } catch (err) { continue; }

    var novas = [];
    fotos.forEach(function (f) {
      var url = String(f.url || "");
      var quebrado = url.indexOf("gettablefileurl") >= 0 || url.indexOf("wixstatic") >= 0;
      if (quebrado) { f.url = ""; limpas++; }
      if (f.url || f.desc) novas.push(f); else removidas++;
    });

    novas.forEach(function (f, k) { f.n = k + 1; });
    valores[i][0] = novas.length ? JSON.stringify(novas) : "";
  }

  intervalo.setValues(valores);
  Logger.log("Endereços quebrados limpos: " + limpas + " | entradas removidas: " + removidas);
}

/** Conta o estado das fotos em toda a folha, sem alterar nada. */
function diagnosticoFotos() {
  var sheet = getSheet();
  var ultima = sheet.getLastRow();
  if (ultima < 2) { Logger.log("Folha vazia."); return; }

  var valores = sheet.getRange(2, 25, ultima - 1, 1).getValues();
  var c = { appsheet: 0, drive: 0, wixstatic: 0, outro: 0, sem_url: 0 };

  valores.forEach(function (v) {
    if (!v[0]) return;
    try {
      JSON.parse(v[0]).forEach(function (f) {
        var u = String(f.url || "");
        if (!u) c.sem_url++;
        else if (u.indexOf("gettablefileurl") >= 0) c.appsheet++;
        else if (u.indexOf("drive.google.com") >= 0) c.drive++;
        else if (u.indexOf("wixstatic") >= 0) c.wixstatic++;
        else c.outro++;
      });
    } catch (err) {}
  });

  Logger.log("Fotos por tipo de endereço: " + JSON.stringify(c));
}

/** ============================ MANUTENÇÃO ============================ **/

/**
 * Acrescenta a coluna "Pasta Drive (URL)" e corrige o nome da coluna do logótipo,
 * SEM apagar dados. Correr uma vez ao passar da v2 para a v3.
 */
function atualizarCabecalho() {
  var sheet = getSheet();
  var colsAtuais = sheet.getLastColumn();

  if (colsAtuais < HEADERS.length) {
    sheet.insertColumnsAfter(colsAtuais, HEADERS.length - colsAtuais);
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setBackground("#0076c9");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  sheet.getRange(1, 10, 1, 12).setBackground("#00548f");
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  Logger.log("Cabeçalho atualizado para v3 (" + HEADERS.length + " colunas). Nenhuma linha de dados foi apagada.");
}

// ATENÇÃO: apaga tudo o que estiver na folha. Só para começar do zero.
function setupSheetHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clear();
  sheet.appendRow(HEADERS);
  atualizarCabecalho();
  sheet.autoResizeColumns(1, HEADERS.length);
  Logger.log("Folha reiniciada com o cabeçalho v3.");
}

/** Teste rápido: confirma o acesso à pasta do Drive. */
function testarPastaDrive() {
  var raiz = DriveApp.getFolderById(PASTA_RAIZ_ID);
  Logger.log("Pasta: " + raiz.getName() + " | " + raiz.getUrl());
}
