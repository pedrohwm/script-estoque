function onEdit(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== "Movimentações") return;

  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  if (col === 7 && row > 1) {
    const nomeItem = sheet.getRange(row, 3).getValue();
    const operacao = sheet.getRange(row, 4).getValue();
    const qtdProduzida = range.getValue();
    const colaborador = sheet.getRange(row, 2).getValue();
    const data = sheet.getRange(row, 1).getValue();

    if (operacao !== "Entrada" || qtdProduzida <= 0) return;

    const receitas = {
      "Nomad Pré-Montado": {
        cor: "#66b46c",
        itens: [
          { item: "Carcaça Nomad", qtd: 1 },
          { item: "PG7 com porca", qtd: 1 },
          { item: "Tampa de PG7", qtd: 2 },
          { item: "Porca PG7", qtd: 2 },
          { item: "PG11 com porca", qtd: 1 },
          { item: "Trava Carcaça Nomad", qtd: 4 }
        ]
      },
      "Nomad Soil Pré-Montado": {
        cor: "#81c784",
        itens: [
          { item: "Nomad Pré-Montado", qtd: 1 },
          { item: "PG7 com porca", qtd: 1 }
        ]
      },
      "Nomad Air Pré-Montado": {
        cor: "#4caf50",
        itens: [
          { item: "Nomad Pré-Montado", qtd: 1 },
          { item: "Tampa de PG7", qtd: 1 },
          { item: "Porca PG7", qtd: 1 },
        ]
      },
      "Painel Solar Pronto": {
        cor: "#c8e6c9",
        itens: [
          { item: "Tampa de carcaça", qtd: 1 },
          { item: "Painel Solar", qtd: 1 },
          { item: "Fio Biocolor Paralelo 2x0.75mm (m)", qtd: 0.3 },
          { item: "Perfil de Vedação Tampa de Nomad (m)", qtd: 0.7 }
        ]
      },
      "Sensor de Fluxo Pronto": {
        cor: "#a5d6a7",
        itens: [
          { item: "Sensor de Fluxo", qtd: 1 },
          { item: "Fio Biocolor Paralelo 2x0.75mm (m)", qtd: 3 },
          { item: "Bucha de Redução PVC Roscável de 1' x 3/4'", qtd: 1 } 
        ]
      }
    };

    const receitaSelecionada = receitas[nomeItem];

    if (receitaSelecionada) {
      receitaSelecionada.itens.forEach(comp => {
        let totalGasto = comp.qtd * qtdProduzida;

        sheet.appendRow([
          data, 
          colaborador, 
          comp.item, 
          "Usado em montagem", 
          "Consumo Montagem", 
          "", 
          totalGasto, 
          "Baixa automática: Montagem de " + nomeItem
        ]);

        let ultimaLinha = sheet.getLastRow();
        sheet.getRange(ultimaLinha, 1, 1, 8).setBackground(receitaSelecionada.cor);
      });
      
      ss.toast("Baixa concluída!", "Estoque", 3);
    }
  }
}