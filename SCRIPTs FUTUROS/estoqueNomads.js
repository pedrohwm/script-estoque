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
        cor: "#977cff",
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
        cor: "#aa78fa",
        itens: [
          { item: "Nomad Pré-Montado", qtd: 1 },
          { item: "PG7 com porca", qtd: 1 }
        ]
      },
      "Nomad Soil LoRa Pronto": {
        cor: "#e17dff",
        itens: [
          { item: "Nomad Soil Pré-Montado", qtd: 1 },
          { item: "Painel Solar Pronto", qtd: 1 },
          { item: "Pinos de Carcaça", qtd: 2 },
          { item: "Placa Principal V1 REV5", qtd: 1 },
          { item: "Placa de Baterias V4 REV1", qtd: 1 },
          { item: "Parafuso Auto Atarraxante Phillips 2.9x9.5 Aço Carbono", qtd: 6 },
          { item: "Cabo JST", qtd: 1 },
          { item: "Adesivo", qtd: 1 },
          { item: "Baterias", qtd: 3 },
          { item: "Antena GPRS (Antena móvel MU-00PI Pentaband 820/960MHz e 1.7/2.5GHz)", qtd: 1 },
          { item: "Perfil de Vedação Antenas (m)", qtd: 0.01 }
        ]
      },
      "Nomad Soil GPRS Pronto": {
        cor: "#e17dff",
        itens: [
          { item: "Nomad Soil Pré-Montado", qtd: 1 },
          { item: "Painel Solar Pronto", qtd: 1 },
          { item: "Pinos de Carcaça", qtd: 2 },
          { item: "Placa Principal V1 REV5", qtd: 1 },
          { item: "Placa de Baterias V4 REV1", qtd: 1 },
          { item: "Parafuso Auto Atarraxante Phillips 2.9x9.5 Aço Carbono", qtd: 6 },
          { item: "Cabo JST", qtd: 1 },
          { item: "Adesivo", qtd: 1 },
          { item: "Baterias", qtd: 3 },
          { item: "Antena GPRS (Antena móvel MU-00PI Pentaband 820/960MHz e 1.7/2.5GHz)", qtd: 1 },
          { item: "Modem LTE/GPRS V1", qtd: 1 },
          { item: "Espaçador Latão Sextavado 5mm M3 Macho/Fêmea", qtd: 1 },
          { item: "Porca Sextavada M3", qtd: 1 },
          { item: "Parafuso M3x4mm", qtd: 1 },
          { item: "Perfil de Vedação Antenas (m)", qtd: 0.01 }
        ]
      },
      "Nomad Air Pré-Montado": {
        cor: "#7270ff",
        itens: [
          { item: "Nomad Pré-Montado", qtd: 1 },
          { item: "Tampa de PG7", qtd: 1 },
          { item: "Porca PG7", qtd: 1 }
        ]
      },
      "Nomad Air LoRa Pronto": {
        cor: "#81a1ff",
        itens: [
          { item: "Nomad Air Pré-Montado", qtd: 1 },
          { item: "Painel Solar Pronto", qtd: 1 },
          { item: "Pinos de Carcaça", qtd: 2 },
          { item: "Placa Principal V1 REV5", qtd: 1 },
          { item: "Placa de Baterias V4 REV1", qtd: 1 },
          { item: "Parafuso Auto Atarraxante Phillips 2.9x9.5 Aço Carbono", qtd: 6 },
          { item: "Cabo JST", qtd: 1 },
          { item: "Adesivo", qtd: 1 },
          { item: "Baterias", qtd: 3 },
          { item: "Antena GPS", qtd: 1 },
          { item: "GPS V1", qtd: 1 },
          { item: "Espaçador Latão Sextavado 12mm M3 Macho/Fêmea", qtd: 1 },
          { item: "Porca Sextavada M3", qtd: 1 },
          { item: "Parafuso M3x8mm", qtd: 1 },
          { item: "Perfil de Vedação Antenas (m)", qtd: 0.01 }
        ]
      },
      "Nomad Air GPRS Pronto": {
        cor: "#81a1ff",
        itens: [
          { item: "Nomad Air Pré-Montado", qtd: 1 },
          { item: "Painel Solar Pronto", qtd: 1 },
          { item: "Pinos de Carcaça", qtd: 2 },
          { item: "Placa Principal V1 REV5", qtd: 1 },
          { item: "Placa de Baterias V4 REV1", qtd: 1 },
          { item: "Parafuso Auto Atarraxante Phillips 2.9x9.5 Aço Carbono", qtd: 6 },
          { item: "Cabo JST", qtd: 1 },
          { item: "Adesivo", qtd: 1 },
          { item: "Baterias", qtd: 3 },
          { item: "Antena GPS", qtd: 1 },
          { item: "GPS V1", qtd: 1 },
          { item: "Espaçador Latão Sextavado 12mm M3 Macho/Fêmea", qtd: 1 },
          { item: "Porca Sextavada M3", qtd: 1 },
          { item: "Parafuso M3x8mm", qtd: 1 },
          { item: "Antena GPRS (Antena móvel MU-00PI Pentaband 820/960MHz e 1.7/2.5GHz)", qtd: 1 },
          { item: "Modem LTE/GPRS V1", qtd: 1 },
          { item: "Espaçador Latão Sextavado 5mm M3 Macho/Fêmea", qtd: 1 },
          { item: "Porca Sextavada M3", qtd: 1 },
          { item: "Parafuso M3x4mm", qtd: 1 },
          { item: "Perfil de Vedação Antenas (m)", qtd: 0.01 }
        ]
      },
      "Painel Solar Pronto": {
        cor: "#ffc400",
        itens: [
          { item: "Tampa de carcaça", qtd: 1 },
          { item: "Painel Solar", qtd: 1 },
          { item: "Fio Biocolor Paralelo 2x0.75mm (m)", qtd: 0.3 },
          { item: "Perfil de Vedação Tampa de Nomad (m)", qtd: 0.7 }
        ]
      },
      "Sensor de Fluxo Pronto": {
        cor: "#33ff00",
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