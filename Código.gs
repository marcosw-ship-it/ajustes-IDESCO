var NOME_ABA_FREQ = "FREQUÊNCIA"; 
var NOME_ABA_DB = "BANCO DE DADOS";
var NOME_ABA_GRUPOS = "GRUPOS";
var LINHA_GRUPO_FREQ = 4; var COLUNA_GRUPO_FREQ = 1;
var LINHA_CABECALHO_FREQ = 6; 

// CONFIGURAÇÃO MANUAL DE COLUNAS OCULTAS
var COLUNAS_PARA_ESCONDER = ["OBS", "Comprovante", "CPF"]; 

// ======================================================
// GATILHOS E MENUS (MANTIDOS)
// ======================================================
function onOpen() { SpreadsheetApp.getUi().createMenu('🚀 SISTEMA').addItem('Abrir Painel Lateral', 'abrirSidebar').addToUi(); }
function abrirSidebar() { var html = HtmlService.createTemplateFromFile('JanelaIdentificacao').evaluate().setTitle('Painel de Controle'); SpreadsheetApp.getUi().showSidebar(html); }
function doGet() { return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Sistema de Frequência').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); }

// ======================================================
// RELATÓRIOS E GRÁFICOS (AJUSTADO PARA CARREGAR DADOS)
// ======================================================
function getDadosRelatorios() { return processarDadosCompleto(); }

function processarDadosCompleto() { 
  try { 
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var sD = ss.getSheetByName(NOME_ABA_DB); 
    var sG = ss.getSheetByName(NOME_ABA_GRUPOS); 
    var dados = sD.getDataRange().getValues(); 
    if (dados.length < 2) return { turmas: [], periodosDetectados: [] }; 

    var Metas = PropertiesService.getScriptProperties().getProperty('METAS_RELATORIO') || "{}"; 
    var cabecalhos = dados[0]; 
    var mapCores = {}; 
    var sGv = sG ? sG.getDataRange().getValues() : []; 
    for (var i = 1; i < sGv.length; i++) mapCores[sGv[i][0]] = sG.getRange(i + 1, 2).getBackground(); 

    var vis = getColunasFiltradas(); 
    var colunasVisuais = vis.heads; 

    // Mapeamento de Períodos para os Gráficos (Manhã, Tarde, Noite)
    var configPeriodos = { 
      "MANHA": { chaves: ["MANHÃ", "MANHA", "MORN"], colunas: [] }, 
      "TARDE": { chaves: ["TARDE", "AFTER"], colunas: [] }, 
      "NOITE": { chaves: ["NOITE", "NIGHT"], colunas: [] } 
    }; 

    for (var c = 0; c < vis.indices.length; c++) { 
      var idxDB = vis.indices[c] - 1; 
      var label = String(cabecalhos[idxDB]).toUpperCase(); 
      for (var key in configPeriodos) { 
        var kws = configPeriodos[key].chaves; 
        for (var k=0; k<kws.length; k++) { 
          if (label.includes(kws[k])) { configPeriodos[key].colunas.push(idxDB); break; } 
        } 
      } 
    } 

    var periodosAtivos = []; 
    if (configPeriodos["MANHA"].colunas.length > 0) periodosAtivos.push("MANHA"); 
    if (configPeriodos["TARDE"].colunas.length > 0) periodosAtivos.push("TARDE"); 
    if (configPeriodos["NOITE"].colunas.length > 0) periodosAtivos.push("NOITE"); 

    var turmas = {}; 
    for (var i = 1; i < dados.length; i++) { 
      var r = dados[i]; var t = r[1]; var n = r[2]; 
      if (!t || !n) continue; 

      if (!turmas[t]) { 
        turmas[t] = { 
          nome: t, 
          cor: mapCores[t] || "#000354", 
          totalAlunos: 0, 
          stats: {}, // Variável que alimenta os Gráficos
          colunas: colunasVisuais, 
          alunos: [] 
        }; 
        periodosAtivos.forEach(p => turmas[t].stats[p] = 0); 
      } 

      turmas[t].totalAlunos++; 

      // Lógica de Contagem para os Gráficos
      periodosAtivos.forEach(function(pNome) { 
        var colsDoPeriodo = configPeriodos[pNome].colunas; 
        var presenteEmTudo = (colsDoPeriodo.length > 0); 
        for(var x=0; x < colsDoPeriodo.length; x++) {
          if(r[colsDoPeriodo[x]] !== "PRESENTE") { presenteEmTudo = false; break; }
        }
        if(presenteEmTudo) turmas[t].stats[pNome]++; 
      }); 

      var pVis = vis.indices.map(idx => r[idx-1] === "PRESENTE");
      turmas[t].alunos.push({ 
        nome: String(n), 
        nomeExibicao: String(n), 
        presencas: pVis 
      }); 
    } 

    var listaFinal = []; 
    Object.keys(turmas).sort().forEach(k => { 
      turmas[k].alunos.sort((a, b) => a.nome.localeCompare(b.nome)); 
      listaFinal.push(turmas[k]); 
    }); 

    return { 
      turmas: listaFinal, 
      periodosDetectados: periodosAtivos, 
      metasSalvas: Metas 
    }; 

  } catch (e) { return { erro: true, msg: "Erro Gráficos: " + e.toString() }; } 
}

// ======================================================
// TODAS AS OUTRAS FUNÇÕES (MANTIDAS EXATAMENTE IGUAIS)
// ======================================================
// Função de edição ajustada para usar a POSIÇÃO (Índice) na hora de atualizar
function minhaFuncaoEditar(e) { 
  if (!e) return; 
  var range = e.range; 
  var sheet = range.getSheet(); 
  var aba = sheet.getName(); 
  
  if (aba.indexOf(NOME_ABA_FREQ) !== 0) return; 
  
  var lin = range.getRow(); 
  var col = range.getColumn(); 
  var val = range.getValue() ? range.getValue().toString().trim() : ""; 
  var old = e.oldValue ? e.oldValue.toString().trim() : ""; 
  
  // Se editou o seletor de turma (Célula B4)
  if (lin === LINHA_GRUPO_FREQ && col === COLUNA_GRUPO_FREQ) { 
    carregarAlunos(val, aba); 
    return; 
  } 
  
  // Se marcou presença (Checkboxes)
  if (lin >= 7 && col >= 4) { 
    var g = sheet.getRange(LINHA_GRUPO_FREQ, COLUNA_GRUPO_FREQ).getValue(); 
    var c = sheet.getRange(LINHA_CABECALHO_FREQ, col).getDisplayValue(); 
    var n = sheet.getRange(lin, 2).getValue(); 
    if (g && c && n) { salvarPresencaEspelhada(g, n, c, val, lin - 7); } 
    return; 
  } 
  
  // Se editou o NOME DO ALUNO (Coluna B)
  if (lin >= 7 && col === 2) { 
    var gAt = sheet.getRange(LINHA_GRUPO_FREQ, COLUNA_GRUPO_FREQ).getValue(); 
    
    // Caso 1: Apagou o nome (Excluir)
    if (val === "") { 
      excluirAlunoPorPosicao(gAt, lin - 7); 
      carregarAlunos(gAt, aba); 
      return; 
    } 
    
    var nNov = toTitleCase(val); 
    var cand = buscarCandidatosDetalhados(nNov); 
    
    // Caso 2: Nome já existe no banco (Conflito/Sidebar)
    // O sistema reverte a edição para o nome antigo antes de abrir a janela
    if (cand.length > 0) { 
      // Se for apenas uma correção de caixa alta/baixa do PRÓPRIO nome, deixa passar
      var ehOProprio = (cand.length === 1 && normalizar(cand[0].nome) === normalizar(old));
      
      if (!ehOProprio) {
        range.setValue(old || ""); // Reverte visualmente
        abrirJanelaDecisao(nNov, gAt, cand, old, lin, aba); 
        return; 
      }
    } 
    
    // Caso 3: Renomear ou Novo
    if (old !== "" && old !== val) { 
      // AQUI ESTÁ A CORREÇÃO: Passamos (lin - 7) que é o índice exato na lista
      atualizarNomeMantendoID(gAt, old, nNov, lin - 7); 
      range.setValue(nNov); // Garante que fica bonito
    } else { 
      range.setValue(nNov); 
      cadastrarNovoAluno(gAt, nNov); 
      carregarAlunos(gAt, aba); 
    } 
  }
}

  function salvarPresencaEspelhada(grupo, nome, cabecalho, valor, indiceNaGrade) { var sD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_DB); var dados = sD.getDataRange().getValues(); var colStatus = dados[0].findIndex(h => String(h).toUpperCase() === cabecalho.toUpperCase()) + 1; if (colStatus <= 0) return; var alunosT = []; for(var i=1; i<dados.length; i++) { if(String(dados[i][1]) === String(grupo)) { alunosT.push({ id: dados[i][0], nome: dados[i][2] }); } } alunosT.sort((a, b) => String(a.nome).localeCompare(String(b.nome))); var alvo = alunosT[indiceNaGrade]; if (!alvo) return; var idAlvo = String(alvo.id); var novoStatus = valor ? "PRESENTE" : "FALTOU"; for(var j=1; j<dados.length; j++) { if(String(dados[j][0]) === idAlvo) { sD.getRange(j + 1, colStatus).setValue(novoStatus); } } SpreadsheetApp.flush(); }
  function processarDecisaoFinal(acao, idAlvo, turmaDestino, nomeNovo, nomeAntigo) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sDB = ss.getSheetByName(NOME_ABA_DB); 
  var sheetAt = ss.getActiveSheet(); 
  
  // Pegamos os dados frescos do banco
  var dDB = sDB.getDataRange().getValues();

  if (acao === "FUNDIR") { 
    // LÓGICA REFORÇADA: Buscar a 'vítima' (quem vai sumir) garantindo que NÃO seja o 'alvo' (quem fica)
    var idVitima = null;
    var nomeBusca = normalizar(nomeAntigo || nomeNovo); // Busca pelo nome original que estava na célula
    
    for(var i=1; i<dDB.length; i++) {
      // Procura: Mesma Turma E Mesmo Nome E ID Diferente do escolhido no painel
      if (String(dDB[i][1]) === turmaDestino && 
          normalizar(dDB[i][2]) === nomeBusca && 
          String(dDB[i][0]) !== String(idAlvo)) {
        idVitima = dDB[i][0];
        break; 
      }
    }
    
    // Se achou um ID diferente com o mesmo nome (ou nome antigo), realiza a fusão
    if (idVitima && idAlvo) {
      fundirPresencas(idAlvo, idVitima); 
    }
    
  } else if (acao === "NOVO") { 
    cadastrarNovoAluno(turmaDestino, nomeNovo); 
  } else { 
    // Lógica para MOVER ou REVISAR (mantida, mas usando loop seguro)
    for(var i=1; i<dDB.length; i++) { 
      if(String(dDB[i][0]) === String(idAlvo)) { 
        if (acao === "MOVER") { 
          sDB.getRange(i+1, 2).setValue(turmaDestino); 
          sDB.getRange(i+1, 3).setValue(nomeNovo); 
        } 
        if (acao === "REVISAR") { 
          var nR = [...dDB[i]]; 
          nR[1]=turmaDestino; 
          nR[2]=nomeNovo; 
          sDB.appendRow(nR); 
        } 
        break; 
      } 
    } 
  } 
  
  SpreadsheetApp.flush(); 
  carregarAlunos(turmaDestino, sheetAt.getName()); 
  limparConflitoResolvido(); 
  return { erro: false }; 
}

function carregarAlunos(grupo, nomeAba) { var ss = SpreadsheetApp.getActiveSpreadsheet(); var sF = ss.getSheetByName(nomeAba || NOME_ABA_FREQ); var sD = ss.getSheetByName(NOME_ABA_DB); if(!sF) return; sF.getRange("B7:B1000").clearContent().setBackground(null).setFontColor("black"); sF.getRange("D7:ZZ1000").removeCheckboxes().clearContent().setBackground(null); if(!grupo) return; var vis = getColunasFiltradas(); if(vis.heads.length > 0) sF.getRange(6, 4, 1, vis.heads.length).setValues([vis.heads.map(h => h.toUpperCase())]).setFontWeight("bold").setBackground("#000354").setFontColor("white").setHorizontalAlignment("center"); var d = sD.getDataRange().getValues(); var filt = d.filter(r => String(r[1]) == grupo).sort((a,b) => String(a[2]).localeCompare(String(b[2]))); if(filt.length > 0) { sF.getRange(7, 2, filt.length, 1).setValues(filt.map(r => [r[2]])); var matrix = filt.map(r => vis.indices.map(idx => r[idx-1] === "PRESENTE")); sF.getRange(7, 4, matrix.length, matrix[0].length).insertCheckboxes().setValues(matrix); var idsVistos = {}; d.forEach(r => { var k=String(r[0]); idsVistos[k]=(idsVistos[k]||0)+1; }); filt.forEach((r, i) => { if(idsVistos[String(r[0])] > 1) sF.getRange(i+7, 2).setBackground("#ffff99"); }); } SpreadsheetApp.flush(); }

function carregarAlunos(grupo, nomeAba) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sF = ss.getSheetByName(nomeAba || NOME_ABA_FREQ); 
  var sD = ss.getSheetByName(NOME_ABA_DB); 
  var sG = ss.getSheetByName(NOME_ABA_GRUPOS); 
  
  if(!sF) return; 
  
  // Limpa a área de alunos
  sF.getRange("B7:B1000").clearContent().setBackground(null).setFontColor("black"); 
  sF.getRange("D7:ZZ1000").removeCheckboxes().clearContent().setBackground(null); 
  
  if(!grupo) {
     sF.getRange(LINHA_GRUPO_FREQ, COLUNA_GRUPO_FREQ).setBackground("white").setFontColor("black");
     return;
  }

  // --- BUSCA A COR DA TURMA ---
  var corGrupo = "#000354"; // Padrão
  if (sG) {
    var dadosG = sG.getDataRange().getValues();
    for (var i = 1; i < dadosG.length; i++) {
      if (String(dadosG[i][0]).toUpperCase().trim() === String(grupo).toUpperCase().trim()) {
        corGrupo = sG.getRange(i + 1, 2).getBackground(); 
        break;
      }
    }
  }
  
  // --- CÁLCULO INTELIGENTE DE CONTRASTE ---
  // Função que decide se a fonte deve ser PRETA ou BRANCA baseada no fundo
  function getCorTexto(hex) {
    if (!hex || hex === "#ffffff" || hex === "white") return "black";
    // Converte Hex para RGB
    var r = parseInt(hex.substr(1, 2), 16);
    var g = parseInt(hex.substr(3, 2), 16);
    var b = parseInt(hex.substr(5, 2), 16);
    // Fórmula de luminosidade (YIQ)
    var yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
    return (yiq >= 128) ? "black" : "white"; // Se for claro (>128), texto preto. Senão, branco.
  }

  var corFonte = getCorTexto(corGrupo);
  // ----------------------------------------

  // APLICA A COR DE FUNDO E A FONTE CALCULADA
  sF.getRange(LINHA_GRUPO_FREQ, COLUNA_GRUPO_FREQ)
    .setBackground(corGrupo)
    .setFontColor(corFonte) // <--- Aplica Preto ou Branco aqui
    .setFontWeight("bold");

  var vis = getColunasFiltradas(); 
  if(vis.heads.length > 0) {
    sF.getRange(6, 4, 1, vis.heads.length)
      .setValues([vis.heads.map(h => h.toUpperCase())])
      .setFontWeight("bold")
      .setBackground("#000354") 
      .setFontColor("white")
      .setHorizontalAlignment("center");
  }

  var d = sD.getDataRange().getValues(); 
  var filt = d.filter(r => String(r[1]) == grupo).sort((a,b) => String(a[2]).localeCompare(String(b[2]))); 
  
  if(filt.length > 0) { 
    sF.getRange(7, 2, filt.length, 1).setValues(filt.map(r => [r[2]])); 
    var matrix = filt.map(r => vis.indices.map(idx => r[idx-1] === "PRESENTE")); 
    sF.getRange(7, 4, matrix.length, matrix[0].length).insertCheckboxes().setValues(matrix); 
    
    var idsVistos = {}; 
    d.forEach(r => { var k=String(r[0]); idsVistos[k]=(idsVistos[k]||0)+1; }); 
    
    filt.forEach((r, i) => { 
      if(idsVistos[String(r[0])] > 1) sF.getRange(i+7, 2).setBackground("#ffff99"); 
    }); 
  } 
  
  SpreadsheetApp.flush(); 
}

function getColunasFiltradas() { 
  var sD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_DB); 
  
  // Pega todos os cabeçalhos da linha 1
  var heads = sD.getRange(1, 1, 1, sD.getLastColumn()).getValues()[0]; 
  var vH = [], vI = []; 

  // Prepara a lista de exclusão: Tira espaços e põe tudo em MAIÚSCULO para garantir
  var listaBloqueio = COLUNAS_PARA_ESCONDER.map(function(nome) {
    return String(nome).toUpperCase().trim();
  });

  // Começa a ler da coluna 4 (Índice 4 = Coluna E, pois 0=A, 1=B, 2=C, 3=D são fixas)
  for (var c = 4; c < heads.length; c++) { 
    var nomeReal = String(heads[c]);
    var nomeComparacao = nomeReal.toUpperCase().trim(); // Normaliza o nome do cabeçalho
    
    // Só adiciona se o nome normalizado NÃO estiver na lista de bloqueio
    if (listaBloqueio.indexOf(nomeComparacao) === -1 && nomeReal.trim() !== "") { 
      vH.push(nomeReal.trim()); 
      vI.push(c + 1); 
    } 
  } 
  
  return { heads: vH, indices: vI }; 
}

function carregarConfigServer() { return PropertiesService.getScriptProperties().getProperty('APP_CONFIG_GLOBAL') || "{}"; }
function salvarConfigServer(json) { PropertiesService.getScriptProperties().setProperty('APP_CONFIG_GLOBAL', json); }
function getPeriodosDisponiveis() { return getColunasFiltradas().heads; }
function processarPresencaWeb(id, periodoNome) { 
  var sDB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_DB); 
  var dados = sDB.getDataRange().getValues(); 
  var colIdx = dados[0].findIndex(h => String(h).trim() === String(periodoNome).trim()) + 1; 
  if (colIdx <= 0) return { erro: true, msg: "Período não encontrado" }; 
  
  // A MÁGICA ACONTECE AQUI: Em vez de ===, ele usa a função que entende a vírgula
  var idxA = dados.findIndex(r => checarId(r[0], id)); 
  
  if (idxA !== -1) { 
    var idAlvo = String(dados[idxA][0]); // Pega todos os IDs juntos (Ex: "123, 456")
    
    // Marca a presença espelhando para todas as turmas
    for(var i=1; i<dados.length; i++) { 
      if(String(dados[i][0]) === idAlvo) sDB.getRange(i+1, colIdx).setValue("PRESENTE"); 
    } 
    SpreadsheetApp.flush(); 
    
    var al = dados[idxA]; 
    var cor = "#333"; 
    var sG = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_GRUPOS); 
    if (sG) { 
      var dG = sG.getDataRange().getValues(); 
      var fG = dG.find(r => String(r[0]).toUpperCase().trim() == String(al[1]).toUpperCase().trim()); 
      if (fG) cor = sG.getRange(dG.indexOf(fG) + 1, 2).getBackground(); 
    } 
    return { erro: false, nome: al[2], grupo: al[1], cor: cor }; 
  } 
  return { erro: true, msg: "ID não localizado" }; 
}
function getPropKey() { var aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(); return 'ABA_KEY_' + aba.replace(/[^a-zA-Z0-9]/g, ""); }
function abrirJanelaDecisao(nome, turma, lista, nomeAntigo, linhaAlvo, nomeAba) { var dados = { nomeBusca: nome, turmaNova: turma, candidatos: lista, nomeAntigo: nomeAntigo, linhaDestino: linhaAlvo, timestamp: new Date().getTime() }; CacheService.getScriptCache().put('ABA_KEY_' + nomeAba.replace(/[^a-zA-Z0-9]/g, ""), JSON.stringify(dados), 600); }
function getDadosSeletor() { return CacheService.getScriptCache().get(getPropKey()); }
function limparConflitoResolvido() { CacheService.getScriptCache().remove(getPropKey()); }
function limparCelulaConflito(lin) { SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(lin, 2).setValue(""); }

// 1. Normalização "Agressiva" (ignora pontos, traços e foca só em letras/números)
function normalizar(t) { 
  if (t === null || t === undefined) return "";
  return String(t).toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Remove acentos
    .replace(/[^A-Z0-9]/g, "") // Remove pontos, traços, espaços (foca no conteúdo)
    .trim();
}

// 2. Formatação Segura (Evita erro com números e o problema do ".0")
function toTitleCase(s){ 
  if (s === null || s === undefined) return "";
  return String(s).toLowerCase().replace(/(?:^|\s)\S/g, a => a.toUpperCase());
}

function cadastrarNovoAluno(g,n){ SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_DB).appendRow([new Date().getTime(),g,n,""]); SpreadsheetApp.flush(); }

// 3. Busca Detalhada Reforçada
function buscarCandidatosDetalhados(n) { 
  var sDB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_DB);
  var d = sDB.getDataRange().getValues(); 
  var b = normalizar(n); 
  var m = []; 
  
  if (b === "") return []; 

  for(var i=1; i<d.length; i++) {
    // Compara a "essência" dos nomes (sem pontos, acentos ou espaços)
    if(normalizar(d[i][2]) === b) {
      m.push({
        id: d[i][0], 
        nome: d[i][2], 
        cpf: d[i][3], 
        turmas: [d[i][1]] 
      });
    }
  }
  return m; 
}

function fundirPresencas(idS, idV) { var sD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_DB); var d = sD.getDataRange().getValues(); var sRows = [], vRows = []; for(var i=1; i<d.length; i++) { if(String(d[i][0])==String(idS)) sRows.push(i); if(String(d[i][0])==String(idV)) vRows.push(i); } for(var c=4; c<d[0].length; c++) { if(vRows.some(idx => d[idx][c]==="PRESENTE") || sRows.some(idx => d[idx][c]==="PRESENTE")) sRows.forEach(idx => sD.getRange(idx+1, c+1).setValue("PRESENTE")); } vRows.reverse().forEach(idx => sD.deleteRow(idx+1)); }
function obterIdPorNomeETurma(g, n) { var d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_DB).getDataRange().getValues(); var f = d.find(r => String(r[1])==g && normalizar(r[2])==normalizar(n)); return f ? f[0] : null; }

// Função reescrita para atualizar pelo ÍNDICE, ignorando duplicatas de nomes
function atualizarNomeMantendoID(grupo, nomeAntigo, nomeNovo, indiceNaLista) { 
  var sD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_DB); 
  var d = sD.getDataRange().getValues(); 
  
  // Reconstrói a lista exatamente como ela aparece na Frequência (Filtrada e Ordenada)
  var listaTurma = [];
  for(var i=1; i<d.length; i++) {
    if(String(d[i][1]) === String(grupo)) {
      listaTurma.push({ dbIndex: i, nome: d[i][2], id: d[i][0] });
    }
  }
  
  // Ordena alfabeticamente (igual ao carregarAlunos)
  listaTurma.sort((a, b) => String(a.nome).localeCompare(String(b.nome)));
  
  // Pega o registro exato baseado na linha que você editou
  var alvo = listaTurma[indiceNaLista];
  
  if(alvo) {
    // Atualiza direto no índice da linha do Banco de Dados
    sD.getRange(alvo.dbIndex + 1, 3).setValue(nomeNovo);
    SpreadsheetApp.flush(); 
  }
}

function excluirAlunoPorPosicao(g, idx) { var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA_DB); var d = s.getDataRange().getValues(); var t = d.map((r,i) => ({r:r, idx:i+1})).filter(x => String(x.r[1]) == g).sort((a,b) => String(a.r[2]).localeCompare(String(b.r[2]))); if (t[idx]) s.deleteRow(t[idx].idx);}

// FUNÇÃO NOVA: Ensina o sistema a procurar IDs separados por vírgula
function checarId(idDaCelula, idProcurado) {
  if (!idDaCelula || !idProcurado) return false;
  // Quebra os IDs pela vírgula e tira os espaços em branco
  var listaIds = String(idDaCelula).split(',').map(function(item) { return item.trim(); });
  // Verifica se o ID que a câmera leu está dentro dessa lista
  return listaIds.indexOf(String(idProcurado).trim()) !== -1;
}
