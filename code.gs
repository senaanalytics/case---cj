// 
// CRIANDO MENU
// NÃO ESQUECER FUNÇAÕ DE ALERTAS
// 



function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CASE- CJ')
    .addItem('Executar validação e limpeza', 'validarDados')
    .addItem('Atualizar situação processos', 'atualizarSituacaoProcesso')
    .addItem('Gerar relatório resumido', 'gerarRelatorio')
    .addSeparator()
    .addItem('Gerar Alertas Críticos', 'gerarAlertasCriticos')
    .addToUi();
}

// 
// Resolvendo erro e setando o nome da aba
// 
function getAbaDados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dados");

  if (!sheet) {
    SpreadsheetApp.getUi().alert("A aba 'Dados' não foi encontrada.");
    return null;
  }

  return sheet;
}

// 
// VALIDAÇÃO
// COLOCAR A LINHA INTEIRA VERMELHA
// 

function validarDados() {
  const sheet = getAbaDados();
  if (!sheet) return;

  const dados = sheet.getDataRange().getValues();
  const cabecalho = dados[0];

  const colunasObrigatorias = [
    "Nome Cliente",
    "Tipo Cidadania",
    "Data Início Processo",
    "Status Processo"
  ];

  const indices = {};
  cabecalho.forEach((coluna, index) => indices[coluna] = index);

  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let linhaRange = sheet.getRange(i + 1, 1, 1, cabecalho.length);
    let dadoAusente = false;

    colunasObrigatorias.forEach(coluna => {
      if (!linha[indices[coluna]]) {
        dadoAusente = true;
        sheet.getRange(i + 1, indices[coluna] + 1)
          .setComment("Dado Ausente");
      }
    });

    if (dadoAusente) {
      linhaRange.setBackground("#ffcccc");
    }

    ["Data Início Processo", "Data Última Atualização", "Data Prevista Conclusão"]
      .forEach(coluna => {
        let valor = linha[indices[coluna]];
        if (valor && !/^\d{4}-\d{2}-\d{2}$/.test(valor)) {
          sheet.getRange(i + 1, indices[coluna] + 1)
            .setBackground("yellow")
            .setComment("Formato de Data Inválido");
        }
      });
  }
}

// 
// ATUALIZAR SITUAÇÃO PROCESSO
// 
function atualizarSituacaoProcesso() {
  const sheet = getAbaDados();
  if (!sheet) return;

  const dados = sheet.getDataRange().getValues();
  const cabecalho = dados[0];

  const indices = {};
  cabecalho.forEach((coluna, index) => indices[coluna] = index);

  let hoje = new Date();
  let colunaSituacao = indices["Situação Processo"];

  if (colunaSituacao === undefined) {
    colunaSituacao = cabecalho.length;
    sheet.getRange(1, colunaSituacao + 1)
      .setValue("Situação Processo");
  }

  for (let i = 1; i < dados.length; i++) {
    let status = dados[i][indices["Status Processo"]] || "";
    let dataPrevista = new Date(dados[i][indices["Data Prevista Conclusão"]]);
    let situacao = "Sem informação";

    if (status === "Concluído") {
      situacao = "Concluído";
    } else if (status === "Cancelado") {
      situacao = "Cancelado";
    } else if (!isNaN(dataPrevista)) {
      let diferencaDias = (dataPrevista - hoje) / (1000 * 60 * 60 * 24);

      if (dataPrevista < hoje) {
        situacao = "Atrasado";
      } else if (diferencaDias <= 30) {
        situacao = "Próximo do Prazo";
      } else {
        situacao = "Em Andamento";
      }
    }

    sheet.getRange(i + 1, colunaSituacao + 1)
      .setValue(situacao);
  }
}

// 
// CRIANDO O RELATÓRIO RESUMIDO
//

function gerarRelatorio() {
  const sheet = getAbaDados();
  if (!sheet) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dados = sheet.getDataRange().getValues();
  const cabecalho = dados[0];

  const indices = {};
  cabecalho.forEach((coluna, index) => {
    indices[coluna] = index;
  });

  let tipoCount = {};
  let statusCount = {};
  let situacaoCount = {};
  let somaConcluidos = 0;
  let qtdConcluidos = 0;
  let pendentes = [];

  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];

    let tipo = linha[indices["Tipo Cidadania"]] || "Sem informação";
    let status = linha[indices["Status Processo"]] || "Sem informação";
    let situacao = linha[indices["Situação Processo"]] || "Sem informação";
    let valor = Number(linha[indices["Valor Serviço (R$)"]]) || 0;
    let pagamento = linha[indices["Status Pagamento"]] || "Sem informação";

    tipoCount[tipo] = (tipoCount[tipo] || 0) + 1;
    statusCount[status] = (statusCount[status] || 0) + 1;
    situacaoCount[situacao] = (situacaoCount[situacao] || 0) + 1;

    if (status === "Concluído") {
      somaConcluidos += valor;
      qtdConcluidos++;
    }

    if (pagamento === "Pendente" || pagamento === "Atrasado") {
      pendentes.push([
        linha[indices["ID Cliente"]] || "Sem informação",
        linha[indices["Nome Cliente"]] || "Sem informação",
        valor
      ]);
    }
  }

  let media = qtdConcluidos ? somaConcluidos / qtdConcluidos : 0;

  let relatorioExistente = ss.getSheetByName("Relatório Resumido");
  if (relatorioExistente) ss.deleteSheet(relatorioExistente);

  let relatorio = ss.insertSheet("Relatório Resumido");

  relatorio.appendRow(["Número de Processos por Tipo Cidadania"]);
  relatorio.appendRow(["Tipo Cidadania", "Quantidade"]);
  for (let key in tipoCount)
    relatorio.appendRow([key, tipoCount[key]]);

  relatorio.appendRow([]);

  relatorio.appendRow(["Número de Processos por Status Processo"]);
  relatorio.appendRow(["Status Processo", "Quantidade"]);
  for (let key in statusCount)
    relatorio.appendRow([key, statusCount[key]]);

  relatorio.appendRow([]);

  relatorio.appendRow(["Número de Processos por Situação Processo"]);
  relatorio.appendRow(["Situação Processo", "Quantidade"]);
  for (let key in situacaoCount)
    relatorio.appendRow([key, situacaoCount[key]]);

  relatorio.appendRow([]);

  relatorio.appendRow(["Média do Valor Serviço (R$) - Processos Concluídos", media]);

  relatorio.appendRow([]);

  relatorio.appendRow(["Top 5 Clientes com Pagamento Pendente ou Atrasado"]);
  relatorio.appendRow(["ID Cliente", "Nome Cliente", "Valor Serviço (R$)"]);

  if (pendentes.length === 0) {
    relatorio.appendRow(["Sem informação", "Sem informação", "Sem informação"]);
  } else {
    pendentes.slice(0, 5).forEach(linha => {
      relatorio.appendRow(linha);
    });
  }
}

// 
// ALERTAS CRÍTICOS
// 


function gerarAlertasCriticos() {
  const sheet = getAbaDados();
  if (!sheet) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dados = sheet.getDataRange().getValues();
  const cabecalho = dados[0];

  const indices = {};
  cabecalho.forEach((coluna, index) => indices[coluna] = index);

  let abaExistente = ss.getSheetByName("⚠️ Alertas Críticos");
  if (abaExistente) ss.deleteSheet(abaExistente);

  let alertas = ss.insertSheet("⚠️ Alertas Críticos");
  alertas.appendRow(cabecalho);

  for (let i = 1; i < dados.length; i++) {
    let situacao = dados[i][indices["Situação Processo"]];
    let pagamento = dados[i][indices["Status Pagamento"]];

    if (situacao === "Atrasado" || pagamento === "Atrasado") {
      alertas.appendRow(dados[i]);
    }
  }
}



/* Atualizar a função de relatório resumido para quando já tiver uma aba com esse nome - *
Atualizar o relatório para quando a célula estiver vazia constar 'Sem informação' na aba resultante - 10/02/2026

Criar alguma função que não está no case. Decidir entre - 'Alerta crítico' - 'Dashboard' - 'Renda total'


Entregar código - 11/02/2026 as 21h00


 */
