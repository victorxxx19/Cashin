// 1. FUNÇÃO PARA DESCOBRIR O NOME
function getNomeUsuarioSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let abaConfig = ss.getSheetByName("BD_Config");
  
  // Se a aba não existir (caso o usuário apague), cria na hora
  if (!abaConfig) {
    abaConfig = ss.insertSheet("BD_Config");
    abaConfig.appendRow(["Chave", "Valor"]);
    abaConfig.appendRow(["NOME_USUARIO", ""]);
  }

  // Tenta ler o nome salvo na célula B2
  const nomeSalvo = abaConfig.getRange("B2").getValue();
  
  if (nomeSalvo && String(nomeSalvo).trim() !== "") {
    return nomeSalvo; // Retorna o nome bonitinho que o usuário escolheu
  }

  // SE NÃO TIVER NOME SALVO: Tenta extrair do e-mail
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      // Pega tudo antes do @ (ex: victor.silva@gmail -> victor.silva)
      let user = email.split('@')[0];
      // Capitaliza a primeira letra (Victor.silva)
      return user.charAt(0).toUpperCase() + user.slice(1);
    }
  } catch (e) {
    return "Visitante"; // Fallback final
  }
  
  return "Visitante";
}

// 2. FUNÇÃO PARA O USUÁRIO ALTERAR O NOME (Pelo Menu)
function configurarNomeUsuario() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Personalização',
    'Como você gostaria de ser chamado no sistema?',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const novoNome = result.getResponseText();
    if(novoNome){
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let abaConfig = ss.getSheetByName("BD_Config");
      if(!abaConfig) abaConfig = ss.insertSheet("BD_Config");
      
      // Garante que a estrutura existe
      if(abaConfig.getLastRow() < 2) {
         abaConfig.getRange("A2").setValue("NOME_USUARIO");
      }
      
      // Salva na B2
      abaConfig.getRange("B2").setValue(novoNome);
      ui.alert(`Pronto! Agora o sistema te chamará de "${novoNome}". Atualize a página do App.`);
    }
  }
}

function abrirGuiaInstalacao() {
  // 1. O TRUQUE DE NOMEAÇÃO (Mantido)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Se ainda estiver com nome de cópia, renomeia
  if (ss.getName().indexOf("Cópia de") > -1) {
     ss.rename("CashIn - App");
  }
  
  // 2. GERA O LINK REAL DO EDITOR
  // Pega o ID único deste script copiado e monta a URL certa
  const scriptId = ScriptApp.getScriptId();
  const urlEditor = `https://script.google.com/home/projects/${scriptId}/edit`;

  // 3. Cria o HTML e injeta a URL
  const template = HtmlService.createTemplateFromFile('GuiaInstalacao');
  template.editorUrl = urlEditor; // <--- Passando a variável para o HTML
  
  const html = template.evaluate()
      .setTitle('Configuração CashIn 🚀');
  
  SpreadsheetApp.getUi().showSidebar(html);
}

// Função para testar se já está configurado (opcional)
function verificarConfiguracao() {
  const email = Session.getActiveUser().getEmail();
  return email;
}

// 3. ATUALIZE SUA FUNÇÃO onOpen PARA TER ESSA OPÇÃO
function onOpen() {
  SpreadsheetApp.getUi().createMenu('CashIn')
    .addItem('Abrir Painel', 'abrirSistema')
    .addSeparator()
    .addItem('⚙️ Alterar meu Nome', 'configurarNomeUsuario') // <--- NOVO
    .addToUi();
}

/* ==================================================
   FUNÇÃO PRINCIPAL: CRIA O WEB APP
   ================================================== */
function doGet() {
  // Cria o HTML a partir do arquivo 'Index'
  var html = HtmlService.createTemplateFromFile('Sistema')
      .evaluate()
      .setTitle('CashIn - Seu dinheiro. Suas Regras.') // O nome que aparece na aba do navegador
      .addMetaTag('viewport', 'width=device-width, initial-scale=1') // Essencial para funcionar no celular
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Permite abrir em iframes se precisar

  return html;
}

/* ==================================================
   FUNÇÃO INCLUDE (Opcional, se quiser separar arquivos no futuro)
   ================================================== */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function abrirSistema() {
  const html = HtmlService.createTemplateFromFile('Sistema')
    .evaluate().setTitle('CashIn').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setWidth(1600).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getDadosDashboard(mes, ano) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaContas = ss.getSheetByName("BD_Contas");
  const abaTrans = ss.getSheetByName("BD_Transacoes");
  const abaCats = ss.getSheetByName("BD_Categorias");
  const abaCartoes = ss.getSheetByName("BD_Cartoes");

  const mTrans = getColMap(abaTrans);
  const mContas = getColMap(abaContas);
  const mCats = abaCats ? getColMap(abaCats) : {};
  const mCartoes = abaCartoes ? getColMap(abaCartoes) : {};

  // 1. Categorias
  let listaCompletaCats = [];
  let categoriasEstruturadas = {};
  if (abaCats) {
      const dadosCats = getDataFromSheet(abaCats);
      listaCompletaCats = dadosCats.map(r => ({
          Tipo: String(r[mCats['Tipo']]).trim(),
          Categoria: String(r[mCats['Categoria']]).trim(),
          Subcategoria: r[mCats['Subcategoria']] ? String(r[mCats['Subcategoria']]).trim() : "",
          Cor: r[mCats['Cor']] || '#94A3B8',
          Icone: r[mCats['Icone']] || 'fas fa-tag'
      }));
      categoriasEstruturadas = processarCategorias(dadosCats, mCats);
  }

  // 2. Contas (Totalizador)
  const rawContas = getDataFromSheet(abaContas);
  let saldoAtualTotal = 0;
  
  const contasFormatadas = rawContas.filter(r => r[mContas['ID_Conta']]).map(r => {
      let valRaw = r[mContas['Saldo_Atual']] !== undefined ? r[mContas['Saldo_Atual']] : r[mContas['Saldo Atual']];
      let saldo = parseMoney(valRaw);
      saldoAtualTotal += saldo; 
      
      return {
          id: r[mContas['ID_Conta']], 
          nome: r[mContas['Nome_Conta']], 
          inst: r[mContas['Instituicao']], 
          saldo: saldo, 
          cor: r[mContas['Cor']]
      };
  });

  // 3. Transações
  const rawTrans = getDataFromSheet(abaTrans);
  
  // Data Limite: Último segundo do mês selecionado
  const dataFimMesSelecionado = new Date(ano, mes + 1, 0, 23, 59, 59);

  let receitasMes = 0, despesasMes = 0, pendenteReceitaMes = 0, pendenteDespesaMes = 0;
  let projecaoReceita = 0, projecaoDespesa = 0;

  const feedTransacoes = [];
  const gastosDetalhados = {};
  const faturasCartoes = {}; 
  const totalUsadoCartoes = {};

  rawTrans.forEach(t => {
    let dataVenc = parseDateSafe(t[mTrans['Data_Vencimento']]);
    if (!dataVenc) return; 

    const valor = parseMoney(t[mTrans['Valor_Parcela']]);
    const tipo = t[mTrans['Tipo']];
    const status = t[mTrans['Status']];
    const idCartao = String(t[mTrans['Cartao_Credito']]); 
    const categoria = t[mTrans['Categoria']] || 'Outros';
    const subcategoria = t[mTrans['Subcategoria']] || 'Geral';

    // A. Mês Selecionado (Cards Visuais do Mês)
    if (dataVenc.getMonth() === mes && dataVenc.getFullYear() === ano) {
       feedTransacoes.push({
          id: t[mTrans['ID_Transacao']], pai: t[mTrans['ID_Pai']], data: dataVenc.toISOString(),
          desc: t[mTrans['Descricao']], valor: valor, tipo: tipo, cat: categoria, subcat: subcategoria,
          tags: t[mTrans['Tags']], obs: t[mTrans['Obs']], status: status, conta: t[mTrans['Conta_Origem']], cartao: idCartao,
          origemNome: '-', parcelaAtual: t[mTrans['Numero_Parcela']], totalParcelas: t[mTrans['Total_Parcelas']]
       });

       if (tipo === 'Receita') {
           receitasMes += valor;
           if(status === 'Pendente') pendenteReceitaMes += valor;
       } 
       else if (tipo === 'Despesa' || tipo === 'Despesa_Cartao') {
           despesasMes += valor;
           if(status === 'Pendente' || status === 'Fatura') pendenteDespesaMes += valor;
           
           // Agrupamento por Categoria
           if (!gastosDetalhados[categoria]) gastosDetalhados[categoria] = { total: 0, subs: {}, cor: '#64748B', icone: 'fas fa-tag' };
           const meta = listaCompletaCats.find(c => c.Categoria === categoria);
           if(meta) { gastosDetalhados[categoria].cor = meta.Cor; gastosDetalhados[categoria].icone = meta.Icone; }
           gastosDetalhados[categoria].total += valor;
           gastosDetalhados[categoria].subs[subcategoria] = (gastosDetalhados[categoria].subs[subcategoria] || 0) + valor;
       }
    }

    // B. Cartões (Acumuladores Gerais)
    if (tipo === 'Despesa_Cartao' && idCartao) {
        // Se não foi pago, abate do limite
        if(status !== 'Pago') totalUsadoCartoes[idCartao] = (totalUsadoCartoes[idCartao] || 0) + valor;
        
        // Se é deste mês, soma na fatura atual visual
        if(dataVenc.getMonth() === mes && dataVenc.getFullYear() === ano) {
            faturasCartoes[idCartao] = (faturasCartoes[idCartao] || 0) + valor;
        }
    }

    // C. CÁLCULO DA PROJEÇÃO
    if (dataVenc <= dataFimMesSelecionado) {
        if (status === 'Pendente' || status === 'Fatura' || status === 'Agendado') {
            if (tipo === 'Receita') {
                projecaoReceita += valor;
            } 
            else if (tipo === 'Despesa' || tipo === 'Despesa_Cartao') {
                projecaoDespesa += valor; 
            }
        }
    }
  });

  feedTransacoes.sort((a, b) => new Date(b.data) - new Date(a.data));

  // --- PROCESSAMENTO FINAL DOS CARTÕES ---
  const rawCartoes = getDataFromSheet(abaCartoes);
  const cartoesFormatados = rawCartoes.filter(r => r[mCartoes['ID_Cartao']]).map(r => {
      const idStr = String(r[mCartoes['ID_Cartao']]);
      const limiteTotal = parseMoney(r[mCartoes['Limite_Total']]);
      const usadoTotal = totalUsadoCartoes[idStr] || 0;
      
      // Cálculo do Limite Disponível (Total - Tudo que não foi pago)
      // Se der negativo, trava em 0
      let disponivel = limiteTotal - usadoTotal;
      
      return {
          id: idStr, 
          nome: r[mCartoes['Nome_Cartao']], 
          inst: r[mCartoes['Instituicao']],
          bandeira: r[mCartoes['Bandeira']], 
          limite: limiteTotal,
          limiteDisponivel: disponivel, // ENVIANDO PRONTO PRO FRONTEND
          fechamento: r[mCartoes['Dia_Fechamento']], 
          vencimento: r[mCartoes['Dia_Vencimento']],
          modoFechamento: r[mCartoes['Modo_Fechamento']] || 'FIXO', 
          contaVinculada: r[mCartoes['Conta Vinculada']],
          faturaAtualValor: faturasCartoes[idStr] || 0,
          totalUsado: usadoTotal,
          faturas: [] // Array vazio para compatibilidade futura
      };
  });

  const balancoMensal = receitasMes - despesasMes; 
  const saldoPrevisto = Number(saldoAtualTotal) + Number(projecaoReceita) - Number(projecaoDespesa);

  return {
    categorias: categoriasEstruturadas,
    listaCompleta: listaCompletaCats,
    nomeUsuario: getNomeUsuarioSistema(),
    resumo: { 
        saldoAtual: saldoAtualTotal, 
        saldoPrevisto: saldoPrevisto, 
        receitas: receitasMes, 
        despesas: despesasMes, 
        balanco: balancoMensal, 
        pendenteReceita: pendenteReceitaMes, 
        pendenteDespesa: pendenteDespesaMes 
    },
    gastosDetalhados: gastosDetalhados,
    contas: contasFormatadas,
    cartoes: cartoesFormatados,
    ultimasTransacoes: feedTransacoes
  };
}

function salvarNomeUsuarioConfig(novoNome) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let abaConfig = ss.getSheetByName("BD_Config");
    
    // Cria se não existir
    if (!abaConfig) {
      abaConfig = ss.insertSheet("BD_Config");
      abaConfig.appendRow(["Chave", "Valor"]);
      abaConfig.appendRow(["NOME_USUARIO", ""]);
    }
    
    // Garante que a célula existe
    if (abaConfig.getLastRow() < 2) {
       abaConfig.getRange("A2").setValue("NOME_USUARIO");
    }
    
    // Salva o valor
    abaConfig.getRange("B2").setValue(novoNome);
    
    return true; // <--- OBRIGATÓRIO RETORNAR ALGO
  } catch(e) {
    throw new Error(e.message);
  }
}

// --- CORREÇÃO: FLUXO DE CAIXA (MODAL) ---
function getFluxoCaixaAnual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaTrans = ss.getSheetByName("BD_Transacoes");
  const abaContas = ss.getSheetByName("BD_Contas");
  
  const m = getColMap(abaTrans);
  const data = getDataFromSheet(abaTrans);
  
  // 1. Saldo Inicial REAL (Soma dos Bancos Hoje)
  let saldoBancosHoje = 0;
  const dadosContas = getDataFromSheet(abaContas);
  const mContas = getColMap(abaContas);
  dadosContas.forEach(r => {
      // Usa parseMoney para garantir que leia certo (ex: 3.201,58)
      let valRaw = r[mContas['Saldo_Atual']] !== undefined ? r[mContas['Saldo_Atual']] : r[mContas['Saldo Atual']];
      saldoBancosHoje += parseMoney(valRaw);
  });

  const hoje = new Date();
  hoje.setHours(0,0,0,0);
  
  // Define o fim do Mês Atual para calcular o "Pendente Geral"
  const fimMesAtual = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0, 23, 59, 59);

  // 2. CÁLCULO DE PENDÊNCIAS GERAIS (IGUAL AO DASHBOARD)
  // Soma TUDO que está pendente (atrasado + mês atual) para corrigir o ponto de partida.
  let pendenteGeralReceita = 0;
  let pendenteGeralDespesa = 0;

  data.forEach(r => {
      let dt = parseDateSafe(r[m['Data_Vencimento']]);
      if (!dt) return;
      
      const status = r[m['Status']];
      const tipo = r[m['Tipo']];
      const valor = parseMoney(r[m['Valor_Parcela']]);

      // Regra de Ouro: Se vence até o fim deste mês E não foi pago, entra na conta do Saldo Previsto
      if (dt <= fimMesAtual) {
         if (status === 'Pendente' || status === 'Fatura' || status === 'Agendado') {
             if (tipo === 'Receita') pendenteGeralReceita += valor;
             else if (tipo === 'Despesa' || tipo === 'Despesa_Cartao') pendenteGeralDespesa += valor;
         }
      }
  });

  const projecao = [];
  const mesesNomes = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

  let saldoAcumulado = 0;

  // Loop dos 12 meses
  for (let i = 0; i < 12; i++) {
      let dataRef = new Date(hoje.getFullYear(), hoje.getMonth() + i, 1);
      let mes = dataRef.getMonth();
      let ano = dataRef.getFullYear();
      
      let totalReceitasMes = 0;
      let totalDespesasMes = 0;
      
      // Varre transações para preencher as colunas "Receitas" e "Despesas" (Visual da tabela)
      data.forEach(r => {
          let dt = parseDateSafe(r[m['Data_Vencimento']]);
          if (!dt) return;
          
          if (dt.getMonth() === mes && dt.getFullYear() === ano) {
              let valor = parseMoney(r[m['Valor_Parcela']]);
              const tipo = r[m['Tipo']];
              
              if (tipo === 'Receita') totalReceitasMes += valor;
              else if (tipo === 'Despesa' || tipo === 'Despesa_Cartao') totalDespesasMes += valor;
          }
      });

      let balanco = totalReceitasMes - totalDespesasMes;
      let saldoFinalMes = 0;

      // --- LÓGICA DO SALDO ACUMULADO ---
      if (i === 0) {
          // MÊS ATUAL (DEZEMBRO):
          // A conta é: Dinheiro que tenho HOJE + O que vai entrar - O que vai sair (incluindo atrasados)
          saldoFinalMes = saldoBancosHoje + pendenteGeralReceita - pendenteGeralDespesa;
      } else {
          // MESES FUTUROS (JANEIRO em diante):
          // A conta é: Saldo do mês anterior + Resultado do mês
          saldoFinalMes = saldoAcumulado + balanco;
      }
      
      saldoAcumulado = saldoFinalMes;

      projecao.push({
          mes: mesesNomes[mes],
          ano: ano,
          receitas: totalReceitasMes,
          despesas: totalDespesasMes,
          balanco: balanco,
          saldoFinal: saldoFinalMes
      });
  }
  
  return projecao;
}

function getGraficoFaturasCartao(cartaoId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaTrans = ss.getSheetByName("BD_Transacoes");
  const abaCartoes = ss.getSheetByName("BD_Cartoes");
  
  // 1. Busca a Cor do Cartão
  let corCartao = '#2563EB'; // Azul padrão (fallback)
  if (abaCartoes) {
      const mCart = getColMap(abaCartoes);
      const dataCart = getDataFromSheet(abaCartoes);
      // Procura a linha deste cartão
      const cartaoEncontrado = dataCart.find(r => String(r[mCart['ID_Cartao']]) === String(cartaoId));
      if (cartaoEncontrado) {
          corCartao = cartaoEncontrado[mCart['Cor']] || '#2563EB';
      }
  }

  const m = getColMap(abaTrans);
  const data = getDataFromSheet(abaTrans);
  const faturas = {};
  const nomesMeses = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];

  data.forEach(r => {
    if (String(r[m['Cartao_Credito']]) !== String(cartaoId)) return;
    
    // Parse seguro de valor
    let valor = 0;
    try {
        let vRaw = r[m['Valor_Parcela']];
        if(typeof vRaw === 'number') valor = vRaw;
        else valor = parseFloat(String(vRaw).replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0;
    } catch(e) { valor = 0; }

    if (valor === 0) return;

    let dataVenc = parseDateSafe(r[m['Data_Vencimento']]);
    if (!dataVenc) return;

    const key = `${dataVenc.getFullYear()}-${String(dataVenc.getMonth() + 1).padStart(2, '0')}`;
    if (!faturas[key]) faturas[key] = 0;
    faturas[key] += valor;
  });

  // Janela Temporal
  const hoje = new Date();
  const inicio = new Date(hoje.getFullYear(), hoje.getMonth() - 6, 1);
  const fim = new Date(hoje.getFullYear(), hoje.getMonth() + 12, 1);

  const keyAtual = `${hoje.getFullYear()}-${String(hoje.getMonth() + 1).padStart(2, '0')}`;
  if(!faturas[keyAtual]) faturas[keyAtual] = 0;

  const chaves = Object.keys(faturas).filter(k => {
      const [y, m] = k.split('-').map(Number);
      const d = new Date(y, m-1, 1);
      return d >= inicio && d <= fim;
  }).sort();

  const labels = [];
  const valores = [];
  const dadosData = [];

  chaves.forEach(key => {
     const [ano, mes] = key.split('-').map(Number);
     labels.push(`${nomesMeses[mes-1]}/${String(ano).slice(2)}`);
     valores.push(faturas[key]);
     dadosData.push({ mes: mes-1, ano: ano });
  });

  // Retorna a COR junto com os dados
  return { labels, valores, dadosData, cor: corCartao };
}

// --- CORREÇÃO 3: Função Auxiliar de Categorias (Mais robusta contra espaços) ---
function processarCategorias(rows, map) {
  const cats = { 'Receita': {}, 'Despesa': {}, 'Despesa_Cartao': {}, 'Transferencia': {} };
  
  rows.forEach(r => {
    // .trim() remove espaços acidentais que quebram a lista
    const tipo = String(r[map['Tipo']]).trim(); 
    const cat = String(r[map['Categoria']]).trim();
    const sub = r[map['Subcategoria']] ? String(r[map['Subcategoria']]).trim() : "";

    if (cats[tipo]) {
      if (!cats[tipo][cat]) cats[tipo][cat] = [];
      if (sub && !cats[tipo][cat].includes(sub)) cats[tipo][cat].push(sub);
    }
    
    // Duplica categorias de Despesa para o seletor de Cartão
    if (tipo === 'Despesa') {
       if (!cats['Despesa_Cartao'][cat]) cats['Despesa_Cartao'][cat] = [];
       if (sub && !cats['Despesa_Cartao'][cat].includes(sub)) cats['Despesa_Cartao'][cat].push(sub);
    }
  });
  return cats;
}

function getFaturaDetalhada(cartaoId, mes, ano) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaTrans = ss.getSheetByName("BD_Transacoes");
    const abaCats = ss.getSheetByName("BD_Categorias"); 
    
    const m = getColMap(abaTrans);
    const data = getDataFromSheet(abaTrans);
    
    // Mapeia Cores E ÍCONES
    const mapCatInfo = {};
    if(abaCats) {
        const mC = getColMap(abaCats);
        const dataC = getDataFromSheet(abaCats);
        dataC.forEach(r => {
            const cat = String(r[mC['Categoria']]).trim();
            mapCatInfo[cat] = {
                cor: r[mC['Cor']] || '#64748B',
                icone: r[mC['Icone']] || 'fas fa-tag'
            }; 
        });
    }

    const itens = data.filter(r => {
        // --- CORREÇÃO AQUI: Parser seguro ---
        let dt = parseDateSafe(r[m['Data_Vencimento']]);
        if (!dt) return false;
        
        return String(r[m['Cartao_Credito']]) === String(cartaoId) && 
               dt.getMonth() === mes && 
               dt.getFullYear() === ano;
    }).map(r => {
        let cat = r[m['Categoria']] ? String(r[m['Categoria']]).trim() : 'Outros';
        const info = mapCatInfo[cat] || { cor: '#64748B', icone: 'fas fa-tag' };
        
        // Data da competência também precisa de parse seguro
        let dtComp = parseDateSafe(r[m['Data_Competencia']]);

        return {
            id: r[m['ID_Transacao']],
            desc: r[m['Descricao']],
            valor: r[m['Valor_Parcela']],
            data: dtComp ? dtComp.toISOString() : null,
            cat: cat,
            color: info.cor,
            icone: info.icone,
            subcat: r[m['Subcategoria']] || '',
            parcela: r[m['Total_Parcelas']] > 1 ? `${r[m['Numero_Parcela']]}/${r[m['Total_Parcelas']]}` : 'À vista'
        };
    });
    
    return itens;
}

// --- NOVA FUNÇÃO: BUSCAR TRANSAÇÃO PARA EDIÇÃO ---
function getTransacaoIndividual(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Transacoes");
  const m = getColMap(aba);
  const data = aba.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][m['ID_Transacao']]) === String(id)) {
       const r = data[i];
       // Retorna objeto formatado igual ao do cache
       return {
          id: r[m['ID_Transacao']],
          pai: r[m['ID_Pai']],
          data: r[m['Data_Vencimento']] instanceof Date ? r[m['Data_Vencimento']].toISOString() : null,
          desc: r[m['Descricao']],
          valor: Number(r[m['Valor_Parcela']]),
          tipo: r[m['Tipo']],
          cat: r[m['Categoria']],
          subcat: r[m['Subcategoria']],
          tags: r[m['Tags']],
          obs: r[m['Obs']],
          status: r[m['Status']],
          conta: r[m['Conta_Origem']],
          cartao: r[m['Cartao_Credito']],
          parcelaAtual: r[m['Numero_Parcela']],
          totalParcelas: r[m['Total_Parcelas']],
          valorTotal: r[m['Valor_Total']]
       };
    }
  }
  return null;
}


// --- 5. DETALHES DE PENDÊNCIAS E CONCLUÍDAS (SEPARADOS) ---
function getPendenciasDetalhadas(tipo, mes, ano) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaTrans = ss.getSheetByName("BD_Transacoes");
  const abaCartoes = ss.getSheetByName("BD_Cartoes");
  const m = getColMap(abaTrans);
  const mCart = getColMap(abaCartoes);
  const data = getDataFromSheet(abaTrans);
  
  const pendentes = [];
  const concluidas = [];
  const faturasAgrupadas = {}; // Para agrupar cartão pendente

  // Mapa de Nomes de Cartão
  const mapCartoes = {};
  const dadosCartoes = getDataFromSheet(abaCartoes);
  dadosCartoes.forEach(c => mapCartoes[c[mCart['ID_Cartao']]] = c[mCart['Nome_Cartao']]);

  data.forEach(r => {
      let d = r[m['Data_Vencimento']];
      if (!(d instanceof Date)) return;
      // Filtra pelo mês/ano selecionado
      if (d.getMonth() !== mes || d.getFullYear() !== ano) return;

      const rTipo = r[m['Tipo']];
      const rValor = Number(r[m['Valor_Parcela']]) || 0;
      const rStatus = r[m['Status']];
      
      // Filtro Geral de Tipo (Receita ou Despesa)
      // Nota: Despesa_Cartao conta como "Despesa" neste contexto
      let matchTipo = false;
      if (tipo === 'Receita' && rTipo === 'Receita') matchTipo = true;
      if (tipo === 'Despesa' && (rTipo === 'Despesa' || rTipo === 'Despesa_Cartao')) matchTipo = true;

      if (!matchTipo) return;

      // Objeto da Transação
      const item = { 
         id: r[m['ID_Transacao']], 
         desc: r[m['Descricao']], 
         valor: rValor, 
         data: d.toISOString(), 
         tipo: rTipo, 
         isFatura: false 
      };

      // --- LÓGICA DE SEPARAÇÃO ---
      
      // 1. CONCLUÍDAS (Pago)
      if (rStatus === 'Pago') {
          // Se for item de cartão pago individualmente ou despesa comum paga
          concluidas.push(item);
      } 
      // 2. PENDENTES (Pendente)
      else if (rStatus === 'Pendente') {
          pendentes.push(item);
      }
      // 3. CARTÃO (Status 'Fatura') - Agrupamento
      else if (rTipo === 'Despesa_Cartao' && rStatus === 'Fatura') {
          const idC = r[m['Cartao_Credito']];
          if (!faturasAgrupadas[idC]) {
              faturasAgrupadas[idC] = { 
                  id: idC, nome: mapCartoes[idC] || 'Cartão', valor: 0, data: d 
              };
          }
          faturasAgrupadas[idC].valor += rValor;
      }
  });

  // Adiciona as Faturas Agrupadas (apenas na lista de Pendentes)
  if (tipo === 'Despesa') {
      Object.values(faturasAgrupadas).forEach(f => {
          if (f.valor > 0) {
              pendentes.push({
                  id: f.id, // ID do Cartão
                  desc: `Fatura ${f.nome}`,
                  valor: f.valor,
                  data: f.data.toISOString(),
                  tipo: 'Despesa',
                  isFatura: true // Marca especial
              });
          }
      });
  }

  // Ordena as listas
  pendentes.sort((a,b) => new Date(a.data) - new Date(b.data));
  concluidas.sort((a,b) => new Date(b.data) - new Date(a.data)); // Concluídas: mais recentes primeiro

  return { pendentes, concluidas };
}

function baixarFaturaCartao(idCartao) {
    // Essa função marca TODAS as compras daquele cartão naquele mês como "Pagas" (ou cria um registro de pagamento)
    // Para simplificar: Vamos criar uma transação de SAÍDA (Despesa) no valor da fatura e marcar os itens como "Pagos".
    // Mas, dado o modelo, o mais simples é criar uma "Despesa" nova representando o pagamento da fatura.
    // E atualizar os itens individuais para 'Pago'.
    
    // NOTA: Para implementar isso perfeitamente, precisamos saber o MÊS que está sendo pago.
    // Como simplificação, vou marcar os itens como Pagos.
    
    // Melhor abordagem V1: Apenas marcar os itens da fatura como 'Pago' para liberar limite e sair da pendência.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("BD_Transacoes");
    const m = getColMap(aba);
    const data = aba.getDataRange().getValues();
    
    // Precisamos saber o mês. Vamos assumir o mês atual selecionado no frontend?
    // O ID passado é do cartão. Vamos varrer e pagar tudo que é "Fatura" antigo ou atual.
    // PERIGO: Pode pagar fatura futura.
    
    // AJUSTE SEGURO: Vamos apenas retornar sucesso por enquanto e pedir para o usuário lançar a despesa de pagamento manualmente?
    // Não, o usuário quer automação.
    
    // Vamos pagar itens com vencimento até hoje + 30 dias (fatura atual/passada)
    const hoje = new Date();
    const limite = new Date();
    limite.setDate(limite.getDate() + 45); // Margem de segurança
    
    let count = 0;
    
    for(let i=1; i<data.length; i++) {
        const row = data[i];
        const rCartao = String(row[m['Cartao_Credito']]);
        const rTipo = row[m['Tipo']];
        const rStatus = row[m['Status']];
        const rData = row[m['Data_Vencimento']];
        
        if (rCartao === String(idCartao) && rTipo === 'Despesa_Cartao' && rStatus === 'Fatura' && rData instanceof Date && rData <= limite) {
            // Marca item individual como Pago
            aba.getRange(i+1, m['Status']+1).setValue('Pago');
            count++;
        }
    }
    
    return { sucesso: true, msg: `${count} itens baixados.` };
}


// --- SALVAR TRANSAÇÃO (VERSÃO FINAL CORRIGIDA) ---
function salvarTransacaoComplexa(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Transacoes");
  const mTrans = getColMap(aba);

  try {
    const idPrincipal = Utilities.getUuid();
    
    // 1. Tratamento de Data (Blindado contra fuso horário)
    let dataCompraBase;
    try {
       if (dados.data && dados.data.includes('-')) {
           const p = dados.data.split('-'); 
           // Ano, Mês (0-11), Dia, Hora 12 (meio dia para evitar virada de data)
           dataCompraBase = new Date(p[0], p[1]-1, p[2], 12, 0, 0);
       } else {
           dataCompraBase = new Date();
       }
    } catch(e) { dataCompraBase = new Date(); }

    // 2. Definição de Repetição e Parcelamento
    // O Frontend manda booleans ou strings "true", garantimos a leitura aqui
    const isParcelado = String(dados.parcelado) === 'true';
    const isRecorrente = String(dados.recorrencia) === 'Fixo';
    
    let numeroParcelas = 1;
    let tipoRepeticao = 'Unico';

    if (isParcelado) {
      let p = parseInt(dados.qtdeParcelas);
      // Se for parcelado, no mínimo 2x
      numeroParcelas = (isNaN(p) || p < 2) ? 2 : p;
      tipoRepeticao = 'Parcelado';
    } 
    else if (isRecorrente) {
      let r = parseInt(dados.qtdeRecorrencia);
      // Se for recorrente, usa o número que vier. Se vier vazio/0, assume 2 para não salvar só 1.
      numeroParcelas = (isNaN(r) || r < 2) ? 2 : r; 
      tipoRepeticao = 'Fixo';
    }

    // 3. Cálculo de Valores
    let valorParcela = 0, valorTotal = 0;
    const valorDigitado = parseFloat(dados.valor) || 0;

    if (tipoRepeticao === 'Parcelado') {
        if (dados.unidadeValor === 'TOTAL') {
            valorTotal = valorDigitado;
            valorParcela = valorDigitado / numeroParcelas;
        } else {
            valorParcela = valorDigitado;
            valorTotal = valorDigitado * numeroParcelas;
        }
    } else {
        // Recorrente ou Único: O valor digitado é o valor da parcela/mensalidade
        valorTotal = valorDigitado; // No banco, valor total costuma ser o da parcela para recorrentes
        valorParcela = valorDigitado;
    }

    // 4. Busca dados do cartão (Para cálculo de vencimento)
    let cardData = null;
    if (dados.tipo === 'Despesa_Cartao' && dados.cartao) {
      cardData = getDadosCartaoCompleto(dados.cartao); 
    }

    let totalParaDebitar = 0;
    const novasLinhas = [];
    const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];

    // 5. Loop de Criação das Linhas (Aqui é onde a mágica acontece)
    for (let i = 0; i < numeroParcelas; i++) {
      
      // Data de Referência (Avança mês a mês)
      let dataRef = new Date(dataCompraBase);
      dataRef.setMonth(dataRef.getMonth() + i);

      // Data de Vencimento Inicial (Igual à referência)
      let dataVencimentoReal = new Date(dataRef);
      let mesRefStr = `${dataRef.getMonth() + 1}/${dataRef.getFullYear()}`;

      // --- LÓGICA DE CARTÃO (FIXO vs DINÂMICO/NUBANK) ---
      if (dados.tipo === 'Despesa_Cartao' && cardData) {
         const diaVencimento = cardData.diaV;
         const valorFechamento = cardData.diaF; // Dia ou Gap
         const modo = cardData.modo; // 'FIXO' ou 'DINAMICO'

         // Cria data candidata para o mês da competência
         let candidataVencimento = new Date(dataRef.getFullYear(), dataRef.getMonth(), diaVencimento);

         if (modo === 'DINAMICO') {
             // LÓGICA NUBANK: Fecha X dias antes
             let dataCorte = new Date(candidataVencimento);
             dataCorte.setDate(dataCorte.getDate() - valorFechamento);
             
             // Se a compra/competência caiu DEPOIS ou IGUAL ao corte, joga pra frente
             if (dataRef >= dataCorte) {
                 candidataVencimento.setMonth(candidataVencimento.getMonth() + 1);
             }
             dataVencimentoReal = candidataVencimento;

         } else {
             // LÓGICA PADRÃO: Fecha dia X
             if (dataRef.getDate() >= valorFechamento) {
                 candidataVencimento.setMonth(candidataVencimento.getMonth() + 1);
             }
             dataVencimentoReal = candidataVencimento;
         }
         
         mesRefStr = `${dataVencimentoReal.getMonth() + 1}/${dataVencimentoReal.getFullYear()}`;
      }

      // Status e Pagamento
      let status = "Pendente";
      let dataPagamento = "";

      if (dados.tipo === 'Despesa_Cartao') {
        status = 'Fatura'; 
      } else if (dados.pago === true || String(dados.pago) === 'true') {
        // Se marcou "Pago", apenas a primeira é paga. As futuras (se recorrente) ficam pendentes?
        // Geralmente Recorrência Fixa marca a 1ª como paga e as outras pendentes.
        // Se for Único, marca pago.
        if (i === 0) { 
          status = "Pago";
          dataPagamento = new Date();
          if(dados.conta) totalParaDebitar += valorParcela;
        } else {
          // Para as próximas parcelas/recorrências, deixamos Pendente para controlar depois
          status = "Pendente"; 
        }
      }

      // Descrição (Adiciona contador se for mais de 1)
      let descFinal = dados.descricao;
      if (numeroParcelas > 1) {
          descFinal += tipoRepeticao === 'Parcelado' ? ` (${i+1}/${numeroParcelas})` : ` (${i+1})`; // Recorrente pode ser só (1), (2)...
      }

      const linhaObj = {
        'ID_Transacao': Utilities.getUuid(), 'ID_Pai': idPrincipal, 
        'Data_Competencia': dataRef,
        'Data_Vencimento': dataVencimentoReal, 
        'Data_Pagamento': dataPagamento, 
        'Mes_Ref': mesRefStr,
        'Descricao': descFinal, 'Tipo': dados.tipo, 
        'Valor_Total': valorTotal, 'Valor_Parcela': valorParcela,
        'Numero_Parcela': (i + 1), 'Total_Parcelas': numeroParcelas,
        'Conta_Origem': dados.conta || "", 'Cartao_Credito': dados.cartao || "",
        'Categoria': dados.categoria || "Outros", 'Subcategoria': dados.subcategoria || "", 
        'Tags': dados.tags || "",
        'Status': status, 'Obs': dados.obs || "", 
        'Recorrencia': tipoRepeticao === 'Fixo' ? 'Fixo' : ''
      };
      
      novasLinhas.push(mapObjectToArray(linhaObj, mTrans, headers.length));
    }
    
    // Grava tudo de uma vez (Bulk Insert)
    if(novasLinhas.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);
    }
    
    // Atualiza saldo da conta (apenas o valor da primeira parcela paga)
    if (totalParaDebitar > 0 && dados.conta && dados.tipo !== 'Despesa_Cartao') {
      atualizarSaldoConta(dados.conta, totalParaDebitar, dados.tipo);
    }
    
    return { sucesso: true, linhasCriadas: novasLinhas.length };
    
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

// --- CORREÇÃO DE SALVAMENTO DE CATEGORIA ---

// Esta é a função que o Frontend está chamando. 
// Ela serve de "ponte" para formatar os dados corretamente.
function salvarCategoriaBack(tipo, nome, cor, icone, subs, mode) {
  // Monta o objeto que a função original esperava
  const dadosFormatados = {
     tipo: tipo,
     nome: nome,
     cor: cor,
     icone: icone,
     subcategorias: subs
  };
  
  return salvarNovaCategoria(dadosFormatados);
}

// Mantenha (ou atualize) sua função salvarNovaCategoria para garantir que funcione
function salvarNovaCategoria(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Categorias");
  const m = getColMap(aba);
  const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  
  try {
    const listaSubs = dados.subcategorias || [];
    // Se a lista estiver vazia e for criação de categoria PAI, adiciona vazio para criar a linha
    if (listaSubs.length === 0) listaSubs.push(""); 

    const novasLinhas = [];
    
    listaSubs.forEach(sub => {
       // Verifica se não está tentando salvar string vazia em subcategoria existente
       if(sub.trim() === "" && listaSubs.length > 1) return;

       const obj = {
         'ID_Cat': Utilities.getUuid(),
         'Tipo': dados.tipo,
         'Categoria': dados.nome,
         'Subcategoria': sub,
         'Icone': dados.icone || 'fas fa-tag',
         'Cor': dados.cor || '#64748B'
       };
       novasLinhas.push(mapObjectToArray(obj, m, headers.length));
    });

    if (novasLinhas.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);
    }
    return { sucesso: true };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function salvarNovaConta(dados) {
  return salvarCadastroGenerico("BD_Contas", {
    'ID_Conta': dados.id || Utilities.getUuid(), 'Nome_Conta': dados.nome, 'Instituicao': dados.banco,
    'Tipo': 'Corrente', 'Saldo_Inicial': parseFloat(dados.saldoInicial), 'Saldo_Atual': parseFloat(dados.saldoInicial),
    'Moeda': "BRL", 'Cor': dados.cor, 'Ativo': "Ativo", 'Data_Criacao': new Date()
  });
}

function salvarNovoCartao(dados) {
  return salvarCadastroGenerico("BD_Cartoes", {
    'ID_Cartao': dados.id || Utilities.getUuid(), 'Nome_Cartao': dados.nome, 'Instituicao': dados.banco,
    'Bandeira': dados.bandeira, 'Limite_Total': parseFloat(dados.limite), 'Dia_Fechamento': parseInt(dados.fechamento),
    'Dia_Vencimento': parseInt(dados.vencimento), 'Cor': dados.cor, 'Ativo': "Ativo", 'Conta Vinculada': dados.contaVinculada || ""
  });
}

function salvarTransferencia(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Transacoes");
  const mTrans = getColMap(aba);
  if (dados.contaOrigem === dados.contaDestino) return { sucesso: false, erro: "Contas iguais." };
  
  try {
    const idPrincipal = Utilities.getUuid();
    const valor = parseFloat(dados.valor);
    const parts = dados.data.split('-'); 
    const dataVenc = new Date(parts[0], parts[1]-1, parts[2], 12, 0, 0);
    const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];

    const baseObj = {
        'ID_Pai': idPrincipal, 'Data_Competencia': new Date(), 'Data_Vencimento': dataVenc, 'Data_Pagamento': new Date(),
        'Mes_Ref': `${dataVenc.getMonth() + 1}/${dataVenc.getFullYear()}`, 'Valor_Total': valor, 'Valor_Parcela': valor,
        'Numero_Parcela': 1, 'Total_Parcelas': 1, 'Categoria': 'Transferência', 'Status': 'Pago', 'Obs': dados.obs || ""
    };

    const objSaida = { ...baseObj, 'ID_Transacao': Utilities.getUuid(), 'Descricao': `Transf. Enviada: ${dados.descricao}`, 'Tipo': 'Despesa', 'Conta_Origem': dados.contaOrigem, 'Subcategoria': 'Enviada' };
    const objEntrada = { ...baseObj, 'ID_Transacao': Utilities.getUuid(), 'Descricao': `Transf. Recebida: ${dados.descricao}`, 'Tipo': 'Receita', 'Conta_Origem': dados.contaDestino, 'Subcategoria': 'Recebida' };

    const linhaSaida = mapObjectToArray(objSaida, mTrans, headers.length);
    const linhaEntrada = mapObjectToArray(objEntrada, mTrans, headers.length);
    aba.getRange(aba.getLastRow() + 1, 1, 2, linhaSaida.length).setValues([linhaSaida, linhaEntrada]);

    atualizarSaldoConta(dados.contaOrigem, valor, 'Despesa');
    atualizarSaldoConta(dados.contaDestino, valor, 'Receita');
    return { sucesso: true, linhasCriadas: 2 };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

// --- SUBSTITUIR/ADICIONAR NO BACKEND.GS ---

// 1. FUNÇÃO DE EXCLUSÃO (NOVA)
function excluirTransacao(id, escopo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Transacoes");
  const m = getColMap(aba);
  
  try {
    const data = aba.getDataRange().getValues();
    // Identifica o ID Pai da transação alvo para exclusão em lote
    let idPaiAlvo = null;
    let linhaAlvo = -1;

    // Busca dados da transação principal
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][m['ID_Transacao']]) === String(id)) {
            idPaiAlvo = data[i][m['ID_Pai']];
            linhaAlvo = i + 1;
            break;
        }
    }

    if (linhaAlvo === -1) return { sucesso: false, erro: "Transação não encontrada" };

    const linhasParaDeletar = [];

    if (escopo === 'single' || !idPaiAlvo) {
        // Deleta apenas a linha encontrada
        linhasParaDeletar.push(linhaAlvo);
    } else {
        // Deleta TODAS com o mesmo ID_Pai (Recorrência)
        // Opcional: Se quiser deletar apenas "desta para frente", teria que comparar datas.
        // Aqui estamos deletando a série inteira conforme pedido "em lote".
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][m['ID_Pai']]) === String(idPaiAlvo)) {
                linhasParaDeletar.push(i + 1);
            }
        }
    }

    // Deleta de baixo para cima para não estragar os índices
    linhasParaDeletar.sort((a, b) => b - a);
    linhasParaDeletar.forEach(rowIndex => {
        aba.deleteRow(rowIndex);
    });

    return { sucesso: true };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

// --- SUBSTITUA A FUNÇÃO salvarEdicaoTransacao NO BACKEND.GS ---

function salvarEdicaoTransacao(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Transacoes");
  const m = getColMap(aba);
  
  if (!dados.id) return { sucesso: false, erro: "ID não fornecido." };

  try {
    const data = aba.getDataRange().getValues();
    const linhasParaEditar = [];
    
    // Conversão segura para String para comparação
    const idAlvo = String(dados.id);
    const paiAlvo = dados.pai ? String(dados.pai) : null;
    const escopo = dados.escopo || 'single';

    // Lógica de Seleção de Linhas
    if (escopo === 'single' || !paiAlvo) {
        // Editar apenas UMA
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][m['ID_Transacao']]) === idAlvo) {
                linhasParaEditar.push(i + 1);
                break; // Achou, para
            }
        }
    } else {
        // Editar TODAS da série (Baseado no ID_Pai)
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][m['ID_Pai']]) === paiAlvo) {
                linhasParaEditar.push(i + 1);
            }
        }
    }

    if (linhasParaEditar.length === 0) return { sucesso: false, erro: "Registro não encontrado para edição." };

    // Prepara dados novos
    const novoValor = parseFloat(dados.valor);
    
    // Tratamento de Data
    let novaDataVencimento = null;
    if (dados.data) {
        const parts = dados.data.split('-'); 
        novaDataVencimento = new Date(parts[0], parts[1]-1, parts[2], 12, 0, 0);
    }
    
    // Loop para aplicar edições nas linhas encontradas
    linhasParaEditar.forEach(rowIndex => {
        // 1. Atualiza Campos Comuns (Texto/Categoria/Valor)
        setCell(aba, rowIndex, m['Descricao'], dados.descricao);
        setCell(aba, rowIndex, m['Categoria'], dados.categoria);
        setCell(aba, rowIndex, m['Subcategoria'], dados.subcategoria);
        setCell(aba, rowIndex, m['Tags'], dados.tags);
        setCell(aba, rowIndex, m['Obs'], dados.obs);
        setCell(aba, rowIndex, m['Valor_Parcela'], novoValor);
        
        // 2. Atualiza Conta ou Cartão
        if (dados.tipo !== 'Despesa_Cartao') {
            setCell(aba, rowIndex, m['Conta_Origem'], dados.conta);
        } else {
            setCell(aba, rowIndex, m['Cartao_Credito'], dados.cartao);
        }

        // 3. Atualiza Data e Status (COMPORTAMENTO ESPECÍFICO)
        // Se for 'single', atualiza data e status daquela parcela específica.
        // Se for 'all', NÃO atualizamos a data de todas para a mesma (pois estragaria o parcelamento),
        // a menos que queiramos recalcular tudo (complexo). Por segurança, em lote, mantemos a data original
        // e alteramos apenas os dados cadastrais (descrição, categoria, valor).
        
        if (escopo === 'single') {
             if(novaDataVencimento) setCell(aba, rowIndex, m['Data_Vencimento'], novaDataVencimento);
             
             if (dados.pago) {
                 setCell(aba, rowIndex, m['Status'], 'Pago');
                 setCell(aba, rowIndex, m['Data_Pagamento'], new Date());
             } else {
                 setCell(aba, rowIndex, m['Status'], dados.tipo === 'Despesa_Cartao' ? 'Fatura' : 'Pendente');
                 setCell(aba, rowIndex, m['Data_Pagamento'], "");
             }
        }
    });

    return { sucesso: true };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

// --- HELPERS ---
function getColMap(sheet) {
  if (sheet.getLastColumn() === 0) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[String(h).trim()] = i);
  return map;
}

function mapObjectToArray(obj, map, totalCols) {
  const row = new Array(totalCols).fill(""); 
  for (const [key, value] of Object.entries(obj)) {
    if (map.hasOwnProperty(key)) row[map[key]] = value;
  }
  return row;
}

function getDataFromSheet(sheet) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}

function salvarCadastroGenerico(nomeAba, objDados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(nomeAba);
    const map = getColMap(aba);
    const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const linhaArray = mapObjectToArray(objDados, map, headers.length);
    aba.appendRow(linhaArray);
    return { sucesso: true };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function setCell(sheet, row, colIndex, val) {
  if (colIndex !== undefined) sheet.getRange(row, colIndex + 1).setValue(val);
}

// --- FUNÇÃO PARA EFETIVAR (DAR BAIXA) ---
function efetivarTransacao(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Transacoes");
  const m = getColMap(aba);
  
  try {
    const data = aba.getDataRange().getValues();
    
    // Procura a linha da transação pelo ID
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][m['ID_Transacao']]) === String(id)) {
        
        // 1. Marca como Pago na planilha
        aba.getRange(i + 1, m['Status'] + 1).setValue('Pago');
        aba.getRange(i + 1, m['Data_Pagamento'] + 1).setValue(new Date());
        
        // 2. Atualiza o Saldo da Conta (Se não for Despesa de Cartão)
        const tipo = data[i][m['Tipo']];
        const valor = parseFloat(data[i][m['Valor_Parcela']]);
        const idConta = data[i][m['Conta_Origem']];

        // Se for Receita ou Despesa comum (não cartão), mexe no saldo da conta
        if (tipo !== 'Despesa_Cartao' && idConta) {
           atualizarSaldoConta(idConta, valor, tipo);
        }

        return { sucesso: true };
      }
    }
    return { sucesso: false, erro: "ID não encontrado." };
  } catch(e) {
    return { sucesso: false, erro: e.message };
  }
}

// --- AUXILIAR PARA ATUALIZAR SALDO DA CONTA ---
// (Caso você já tenha essa função, certifique-se que ela está igual a esta)
function atualizarSaldoConta(idConta, valor, tipo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Contas");
  const m = getColMap(aba);
  const data = aba.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][m['ID_Conta']]) === String(idConta)) {
      let saldoAtual = parseFloat(data[i][m['Saldo_Atual']] || 0);
      
      if (tipo === 'Receita') {
        saldoAtual += valor;
      } else if (tipo === 'Despesa' || tipo === 'Transferencia') { 
        // Transferencia enviada conta como despesa na origem
        saldoAtual -= valor;
      }
      
      aba.getRange(i + 1, m['Saldo_Atual'] + 1).setValue(saldoAtual);
      break; 
    }
  }
}

function getDadosCartaoCompleto(cartaoId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Cartoes");
  if (!aba || aba.getLastRow() < 2) return null;
  
  const m = getColMap(aba);
  const dados = aba.getRange(2, 1, aba.getLastRow()-1, aba.getLastColumn()).getValues();
  
  for(let i=0; i<dados.length; i++){
    if(String(dados[i][m['ID_Cartao']]) === String(cartaoId)){
      return {
        diaF: parseInt(dados[i][m['Dia_Fechamento']]) || 1,
        diaV: parseInt(dados[i][m['Dia_Vencimento']]) || 10,
        modo: dados[i][m['Modo_Fechamento']] || 'FIXO' // <--- LÊ O MODO
      };
    }
  }
  return { diaF: 1, diaV: 10, modo: 'FIXO' };
}

function parseMoney(val) {
  if (val === undefined || val === null || val === "") return 0;
  if (typeof val === 'number') return val;
  let clean = val.toString().replace('R$', '').replace(/\./g, '').replace(',', '.').trim();
  return parseFloat(clean) || 0;
}

function getAssetsMap() {
  return { banks: { 'Nubank': { color: '#820AD1', domain: 'nubank.com.br' }, 'Inter': { color: '#FF7A00', domain: 'inter.co' }, 'PicPay': { color: '#11C76F', domain: 'picpay.com' }, 'Mercado Pago': { color: '#009EE3', domain: 'mercadopago.com.br' }, 'Itau': { color: '#EC7000', domain: 'itau.com.br' }, 'Bradesco': { color: '#CC092F', domain: 'bradesco.com.br' }, 'Santander': { color: '#EC0000', domain: 'santander.com.br' }, 'Neon': { color: '#00FFFF', domain: 'neon.com.br' }, 'C6': { color: '#242424', domain: 'c6bank.com.br' }, 'Caixa': { color: '#005CA9', domain: 'caixa.gov.br' }, 'Brasil': { color: '#F8D117', domain: 'bb.com.br' }, 'XP': { color: '#000000', domain: 'xpi.com.br' } } };
}

// --- BACKEND: BUSCAR HISTÓRICO COMPLETO DO CARTÃO ---
function getHistoricoCompletoCartao(cartaoId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Transacoes");
  const m = getColMap(aba);
  const data = getDataFromSheet(aba);
  
  // Objeto para agrupar: { '12/2025': { total: 0, itens: [], sort: 202512 } }
  const faturas = {};

  data.forEach(r => {
    // Verifica se é do cartão solicitado e se tem valor
    if (String(r[m['Cartao_Credito']]) !== String(cartaoId)) return;
    
    const valor = parseFloat(r[m['Valor_Parcela']]) || 0;
    if (valor === 0) return;

    let dataVenc = r[m['Data_Vencimento']];
    if (!(dataVenc instanceof Date)) return; // Ignora se não tiver data válida

    // Cria chave de agrupamento (Mês/Ano)
    const mes = dataVenc.getMonth() + 1;
    const ano = dataVenc.getFullYear();
    const chave = `${mes.toString().padStart(2, '0')}/${ano}`;
    const sortKey = (ano * 100) + mes; // Ex: 202512 para ordenação

    if (!faturas[chave]) {
      faturas[chave] = {
        mesRef: chave,
        mesRefSort: sortKey,
        total: 0,
        itens: []
      };
    }

    faturas[chave].total += valor;
    faturas[chave].itens.push({
      data: dataVenc.toISOString(),
      desc: r[m['Descricao']],
      valor: valor,
      parcela: r[m['Total_Parcelas']] > 1 ? `${r[m['Numero_Parcela']]}/${r[m['Total_Parcelas']]}` : 'À vista'
    });
  });

  // Transforma objeto em array
  const resultado = Object.values(faturas);
  
  // Opcional: Ordenar itens dentro de cada fatura por data
  resultado.forEach(fat => {
    fat.itens.sort((a, b) => new Date(a.data) - new Date(b.data));
  });

  return resultado;
}

// --- ATUALIZAR DADOS DA CONTA ---
function atualizarDadosConta(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Contas");
  const m = getColMap(aba);
  const data = aba.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][m['ID_Conta']]) === String(dados.id)) {
       // Atualiza colunas
       aba.getRange(i + 1, m['Nome_Conta'] + 1).setValue(dados.nome);
       aba.getRange(i + 1, m['Instituicao'] + 1).setValue(dados.banco);
       // Opcional: Atualizar saldo manual (cuidado para não quebrar histórico)
       // aba.getRange(i + 1, m['Saldo_Atual'] + 1).setValue(parseFloat(dados.saldoInicial));
       return { sucesso: true };
    }
  }
  return { sucesso: false, erro: "Conta não encontrada" };
}

// --- ATUALIZAR DADOS DO CARTÃO ---
function atualizarDadosCartao(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Cartoes");
  const m = getColMap(aba);
  const data = aba.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][m['ID_Cartao']]) === String(dados.id)) {
       aba.getRange(i + 1, m['Nome_Cartao'] + 1).setValue(dados.nome);
       aba.getRange(i + 1, m['Instituicao'] + 1).setValue(dados.banco);
       aba.getRange(i + 1, m['Bandeira'] + 1).setValue(dados.bandeira);
       aba.getRange(i + 1, m['Limite_Total'] + 1).setValue(parseFloat(dados.limite));
       aba.getRange(i + 1, m['Dia_Fechamento'] + 1).setValue(parseInt(dados.fechamento));
       aba.getRange(i + 1, m['Dia_Vencimento'] + 1).setValue(parseInt(dados.vencimento));
       aba.getRange(i + 1, m['Conta Vinculada'] + 1).setValue(dados.contaVinculada);
       return { sucesso: true };
    }
  }
  return { sucesso: false, erro: "Cartão não encontrado" };
}

// --- GESTÃO DE METAS (BACKEND) ---

function getMetas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName("BD_Metas");
  
  // Cria a aba se não existir
  if (!aba) {
    aba = ss.insertSheet("BD_Metas");
    // Cabeçalhos
    aba.appendRow(["ID_Meta", "Nome", "Valor_Alvo", "Valor_Atual", "Cor", "Icone", "Data_Limite"]);
    return [];
  }
  
  const m = getColMap(aba);
  const data = getDataFromSheet(aba);
  
  return data.map(r => ({
      id: r[m['ID_Meta']],
      nome: r[m['Nome']],
      alvo: parseFloat(r[m['Valor_Alvo']] || 0),
      atual: parseFloat(r[m['Valor_Atual']] || 0),
      cor: r[m['Cor']] || '#10B981',
      icone: r[m['Icone']] || 'fas fa-trophy',
      dataLimite: r[m['Data_Limite']] ? new Date(r[m['Data_Limite']]).toISOString() : null
  }));
}

function salvarNovaMeta(dados) {
  const obj = {
      'ID_Meta': Utilities.getUuid(),
      'Nome': dados.nome,
      'Valor_Alvo': parseFloat(dados.alvo),
      'Valor_Atual': parseFloat(dados.atual),
      'Cor': dados.cor,
      'Icone': dados.icone,
      'Data_Limite': dados.dataLimite ? new Date(dados.dataLimite) : ""
  };
  
  return salvarCadastroGenerico("BD_Metas", obj);
}

function atualizarSaldoMeta(id, valorAdicional) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Metas");
  const m = getColMap(aba);
  const data = aba.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][m['ID_Meta']]) === String(id)) {
       let atual = parseFloat(data[i][m['Valor_Atual']] || 0);
       let novo = atual + parseFloat(valorAdicional);
       aba.getRange(i + 1, m['Valor_Atual'] + 1).setValue(novo);
       return { sucesso: true, novoSaldo: novo };
    }
  }
  return { sucesso: false, erro: "Meta não encontrada" };
}

// --- NOVO: BUSCAR CLIMA PELO SERVIDOR (CORRIGE O BUG DE CARREGAMENTO) ---
function getClimaServidor() {
  try {
    // Coordenadas de Feira de Santana
    const url = "https://api.open-meteo.com/v1/forecast?latitude=-12.2667&longitude=-38.9667&current=temperature_2m,weather_code&timezone=America%2FSao_Paulo";
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if(data && data.current) {
       return { 
         temp: Math.round(data.current.temperature_2m),
         code: data.current.weather_code
       };
    }
  } catch(e) {
    return null; // Se falhar, segue sem clima
  }
  return null;
}

function FORCAR_CORRECAO_NUBANK() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaTrans = ss.getSheetByName("BD_Transacoes");
  const abaCartoes = ss.getSheetByName("BD_Cartoes");
  
  const mTrans = getColMap(abaTrans);
  const mCart = getColMap(abaCartoes);
  
  // 1. Acha o ID do Nubank
  const dadosCartoes = abaCartoes.getDataRange().getValues();
  let idNubank = null;
  
  for(let i=1; i<dadosCartoes.length; i++) {
     const nome = String(dadosCartoes[i][mCart['Nome_Cartao']]).toLowerCase();
     // Procura por "nubank" no nome (ajuste se o nome for diferente)
     if(nome.includes("nubank")) {
        idNubank = String(dadosCartoes[i][mCart['ID_Cartao']]);
        break;
     }
  }
  
  if(!idNubank) {
     SpreadsheetApp.getUi().alert("Erro: Cartão Nubank não encontrado na aba BD_Cartoes.");
     return;
  }

  // CONFIGURAÇÃO FORÇADA PARA O NUBANK
  const DIA_VENCIMENTO = 6;
  const DIAS_ANTES_FECHAMENTO = 7; 

  const dataTrans = abaTrans.getDataRange().getValues();
  let alterados = 0;

  // 2. Varre todas as transações
  for(let i=1; i<dataTrans.length; i++) {
     const row = dataTrans[i];
     const idCartaoRow = String(row[mTrans['Cartao_Credito']]);
     
     // Verifica se é o cartão Nubank
     if(idCartaoRow === idNubank) {
        
        let dataCompra = row[mTrans['Data_Competencia']];
        if(!(dataCompra instanceof Date)) continue;

        // --- CÁLCULO DA DATA CORRETA ---
        // 1. Define Vencimento Candidato (Dia 06 do mesmo mês da compra)
        // Ex: Compra 01/11 -> Candidato 06/11
        let novoVencimento = new Date(dataCompra.getFullYear(), dataCompra.getMonth(), DIA_VENCIMENTO);
        
        // 2. Define Data de Fechamento (Candidato - 7 dias)
        // Ex: 06/11 - 7 dias = 30/10
        let dataFechamento = new Date(novoVencimento);
        dataFechamento.setDate(dataFechamento.getDate() - DIAS_ANTES_FECHAMENTO);
        
        // 3. Compara: A compra foi DEPOIS ou IGUAL ao fechamento?
        // Ex: 01/11 >= 30/10? SIM.
        if (dataCompra >= dataFechamento) {
            // Então pula para o mês seguinte (Dezembro)
            novoVencimento.setMonth(novoVencimento.getMonth() + 1);
        }
        
        // --- GRAVA NA PLANILHA ---
        // Atualiza Data Vencimento
        abaTrans.getRange(i+1, mTrans['Data_Vencimento']+1).setValue(novoVencimento);
        
        // Atualiza Mes_Ref (Visual) -> "12/2025"
        let mesRef = `${novoVencimento.getMonth()+1}/${novoVencimento.getFullYear()}`;
        abaTrans.getRange(i+1, mTrans['Mes_Ref']+1).setValue(mesRef);
        
        // Atualiza Status para "Fatura" (Garante que apareça nos cálculos)
        if (row[mTrans['Status']] !== 'Pago') {
             abaTrans.getRange(i+1, mTrans['Status']+1).setValue('Fatura');
        }

        alterados++;
     }
  }
  
  SpreadsheetApp.getUi().alert(`Sucesso! ${alterados} transações do Nubank foram corrigidas para o vencimento dia ${DIA_VENCIMENTO} (fechamento ${DIAS_ANTES_FECHAMENTO} dias antes).`);
}

// --- HELPER: CONVERTE QUALQUER COISA PARA DATA ---
function parseDateSafe(v) {
  if (!v) return null;
  if (v instanceof Date) return v; // Já é data
  
  if (typeof v === 'string') {
    v = v.trim();
    // Tenta formato PT-BR (DD/MM/YYYY)
    if (v.includes('/')) {
       const p = v.split('/');
       // p[2]=Ano, p[1]=Mes, p[0]=Dia
       if(p.length === 3) return new Date(p[2], p[1]-1, p[0], 12, 0, 0);
    }
    // Tenta formato ISO (YYYY-MM-DD)
    if (v.includes('-')) {
       const p = v.split('-');
       if(p.length === 3) return new Date(p[0], p[1]-1, p[2], 12, 0, 0);
    }
  }
  return null;
}

// --- AUTOCOMPLETE INTELIGENTE ---
function getSugestoesPreenchimento() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BD_Transacoes");
  const m = getColMap(aba);
  
  // Pega os dados brutos
  const data = getDataFromSheet(aba);
  const mapa = {}; // Chave: Descrição -> Valor: {cat, sub, conta, cartao, valor...}

  // Varre do fim para o começo (para pegar sempre o mais recente)
  for (let i = data.length - 1; i >= 0; i--) {
      const r = data[i];
      const desc = String(r[m['Descricao']]).trim();
      
      // Se tiver descrição e ainda não mapeamos (ou seja, é o mais recente)
      if (desc && !mapa[desc]) {
          mapa[desc] = {
              cat: r[m['Categoria']],
              sub: r[m['Subcategoria']],
              conta: r[m['Conta_Origem']],
              cartao: r[m['Cartao_Credito']],
              valor: r[m['Valor_Parcela']], // Pega o valor da parcela como sugestão
              tags: r[m['Tags']],
              tipo: r[m['Tipo']]
          };
      }
      
      // Limite de segurança: se tiver mais de 500 descrições únicas, para (pra não pesar o front)
      if (Object.keys(mapa).length > 500) break;
  }
  
  return mapa;
}


/******************************************************
 * CONFIGURAÇÃO GLOBAL DA IA
 ******************************************************/
const API_KEY = 'AIzaSyBUYOWF6ADA2I1jotnWUPQx3kwM1SmIS2k'; 

// ✅ CORREÇÃO FINAL: Usando os modelos que SEUS LOGS confirmaram
const MODEL_MAIN = 'models/gemini-2.5-flash'; 
const MODEL_FALLBACK = 'models/gemini-2.5-pro';

const LOG_SHEET_NAME = "LOG_CHAT_IA";

/******************************************************
 * FUNÇÃO DE ENTRADA
 ******************************************************/
function chamarGemini(prompt, historico = []) {
  try {
    // 1. AÇÕES RÁPIDAS
    const intencao = detectarIntencao(prompt);
    const jsonAcao = executarAcaoFinanceira(intencao);
    if (jsonAcao) {
        registrarLog(prompt, jsonAcao, "ACTION");
        return { sucesso: true, texto: jsonAcao };
    }

    // 2. CONTEXTO (Com tratamento de erro para não travar)
    let contexto = {};
    try {
       contexto = getContextoFinanceiroParaIA();
    } catch(e) {
       contexto = { erro: "Não foi possível ler os dados detalhados: " + e.message };
    }

    // 3. GERA RESPOSTA
    const resposta = gerarRespostaUnificada(prompt, historico, contexto);
    
    registrarLog(prompt, resposta, "API");
    return { sucesso: true, texto: resposta };

  } catch (e) {
    return { sucesso: false, erro: "Erro Crítico: " + e.message };
  }
}

/******************************************************
 * GERAÇÃO DE RESPOSTA (BLINDADA)
 ******************************************************/
function gerarRespostaUnificada(prompt, historico, contexto) {
  
  // Formata histórico como texto simples para evitar erros de estrutura JSON da API
  let historicoTexto = "";
  if (historico && historico.length > 0) {
     historicoTexto = historico.map(h => 
       `${h.role === 'user' ? 'USUÁRIO' : 'IA'}: ${h.text}`
     ).join("\n");
  }

  const sistema = `
    VOCÊ É: CashIn Assistant, um consultor financeiro pessoal estratégico.
    
    DADOS DO USUÁRIO (Contexto Real):
    ${JSON.stringify(contexto)}

    HISTÓRICO RECENTE:
    ${historicoTexto}

    PERGUNTA ATUAL:
    "${prompt}"

    DIRETRIZES:
    1. Responda baseando-se ESTRITAMENTE nos dados.
    2. Analise o campo "projecao_futura" para responder sobre 2026.
    3. Se o saldo futuro for negativo, ALERTE.
    4. Seja direto e use Markdown (**negrito**) nos valores.
  `;

  const contents = [{ parts: [{ text: sistema }] }];

  // Tenta 2.5 Flash
  let r1 = chamarModelo(MODEL_MAIN, contents);
  if (r1.sucesso) return r1.texto;

  // Tenta 2.5 Pro
  let r2 = chamarModelo(MODEL_FALLBACK, contents);
  if (r2.sucesso) return r2.texto;

  // Retorna erro detalhado se falhar
  return "Erro de conexão com a IA (Modelos 2.5 não responderam).";
}

function chamarModelo(modelo, contents) {
  const url = `https://generativelanguage.googleapis.com/v1beta/${modelo}:generateContent?key=${API_KEY}`;
  const payload = { 
      contents: contents, 
      generationConfig: { temperature: 0.5, maxOutputTokens: 2000 } 
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
       const json = JSON.parse(response.getContentText());
       if (json.candidates && json.candidates.length > 0) {
          return { sucesso: true, texto: json.candidates[0].content.parts[0].text };
       }
    }
    return { sucesso: false };
  } catch (e) { return { sucesso: false }; }
}

/******************************************************
 * PREPARAÇÃO DE DADOS (COM FUTURO)
 ******************************************************/
function getContextoFinanceiroParaIA() {
  const hoje = new Date();
  const dados = getDadosDashboard(hoje.getMonth(), hoje.getFullYear());
  
  // Pega o futuro para ela saber responder sobre 2026
  const projecao = getFluxoCaixaAnual();
  const resumoFuturo = projecao.map(p => 
    `[${p.mes}/${p.ano}] Saldo: ${parseInt(p.saldoFinal)}`
  ).join(' | ');

  return {
    data_hoje: new Date().toLocaleDateString('pt-BR'),
    situacao_atual: {
       saldo_bancos: dados.resumo.saldoAtual,
       previsao_fechamento_mes: dados.resumo.saldoPrevisto
    },
    projecao_futura: resumoFuturo, // <--- O SEGREDO PARA 2026
    ultimas_transacoes: dados.ultimasTransacoes.slice(0, 8).map(t => 
       `${t.data.substr(0,10)}: ${t.desc} (${parseInt(t.valor)})`
    )
  };
}

// ... (Mantenha as funções detectarIntencao, executarAcaoFinanceira e registrarLog que já estavam certas) ...
function detectarIntencao(texto) {
  if(!texto) return null;
  const t = texto.toLowerCase();
  if (t.includes('cartões') || t.includes('cartoes') || t.includes('fatura')) return 'IR_CARTOES';
  if (t.includes('contas') || t.includes('bancos')) return 'IR_CONTAS';
  if (t.includes('visão geral') || t.includes('dashboard')) return 'IR_DASHBOARD';
  if (t.includes('projeção') || t.includes('futuro') || t.includes('previsão')) return 'ABRIR_PROJECAO';
  if (t.includes('nova meta') || t.includes('criar meta')) return 'ABRIR_META';
  if (t.includes('nova receita') || t.includes('ganhei')) return 'ABRIR_RECEITA';
  if (t.includes('nova despesa') || t.includes('gastei')) return 'ABRIR_DESPESA';
  if (t.includes('buscar') || t.includes('procure por')) return 'BUSCAR_TRANSACAO';
  return null; 
}

function executarAcaoFinanceira(intencao) {
  switch (intencao) {
    case 'IR_CARTOES': return JSON.stringify({ action: "NAVIGATE", target: "view-cartoes", msg: "Abrindo **Cartões**." });
    case 'IR_CONTAS': return JSON.stringify({ action: "NAVIGATE", target: "view-contas", msg: "Abrindo **Contas**." });
    case 'IR_DASHBOARD': return JSON.stringify({ action: "NAVIGATE", target: "view-dashboard", msg: "Voltando ao **Início**." });
    case 'ABRIR_PROJECAO': return JSON.stringify({ action: "OPEN_MODAL", target: "projecao", msg: "Abrindo **Projeção**." });
    case 'ABRIR_META': return JSON.stringify({ action: "OPEN_MODAL", target: "meta", msg: "Criar **Nova Meta**." });
    case 'ABRIR_RECEITA': return JSON.stringify({ action: "OPEN_MODAL", target: "receita", msg: "Nova **Receita**." });
    case 'ABRIR_DESPESA': return JSON.stringify({ action: "OPEN_MODAL", target: "despesa", msg: "Nova **Despesa**." });
    case 'BUSCAR_TRANSACAO': return JSON.stringify({ action: "FILTER_TRANS", value: "", msg: "Busca iniciada." });
  }
  return null;
}

function registrarLog(prompt, resposta, origem) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sheet) { sheet = ss.insertSheet(LOG_SHEET_NAME); sheet.appendRow(["Data", "Prompt", "Resposta", "Origem"]); }
    sheet.appendRow([new Date(), prompt, resposta, origem]);
  } catch(e) {}
}

function processarPagamentoFatura(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaTrans = ss.getSheetByName("BD_Transacoes");
  const abaCartoes = ss.getSheetByName("BD_Cartoes");
  const abaContas = ss.getSheetByName("BD_Contas"); // Necessário para debitar a conta

  if (!abaTrans || !abaCartoes || !abaContas) {
      return { sucesso: false, erro: "Abas de dados (BD_...) não encontradas." };
  }

  try {
    const m = getColMap(abaTrans);
    const dataRange = abaTrans.getDataRange();
    const valores = dataRange.getValues();
    
    // Data do Pagamento (Hoje)
    const dataPag = new Date(dados.data); 
    
    // Mês/Ano da Fatura que estamos pagando (Vindo do front)
    const mesAlvo = parseInt(dados.mesRef);
    const anoAlvo = parseInt(dados.anoRef);
    
    let totalBaixado = 0;
    let countItens = 0;
    
    // 1. VARRER E ATUALIZAR ITENS DA FATURA
    // Ao invés de criar uma linha nova, vamos editar as linhas existentes
    for (let i = 1; i < valores.length; i++) {
        const row = valores[i];
        
        // Pega dados da linha
        const rCartao = String(row[m['Cartao_Credito']]);
        const rStatus = String(row[m['Status']]);
        const rTipo = String(row[m['Tipo']]);
        
        // Parser seguro da data de vencimento da compra
        let rVenc = row[m['Data_Vencimento']];
        if (typeof rVenc === 'string') rVenc = parseDateSafe(rVenc);
        if (!(rVenc instanceof Date)) continue;

        // VERIFICA SE É ITEM DA FATURA SELECIONADA
        // 1. Mesmo Cartão
        // 2. Tipo 'Despesa_Cartao'
        // 3. Status 'Fatura'
        // 4. Mês e Ano de Vencimento batem com a fatura visualizada
        
        const mesmoMesAno = (rVenc.getMonth() === mesAlvo && rVenc.getFullYear() === anoAlvo);
        
        // Opcional: Baixa também atrasados (vencimento anterior e ainda 'Fatura')
        const atrasado = (rVenc < new Date(anoAlvo, mesAlvo, 1)); 

        if (rCartao === String(dados.cartaoId) && 
            rTipo === 'Despesa_Cartao' && 
            rStatus === 'Fatura' && 
            (mesmoMesAno || atrasado)) {
            
            const rValor = parseFloat(row[m['Valor_Parcela']]) || 0;
            
            // --- A MÁGICA: ATUALIZA A LINHA EXISTENTE ---
            
            // 1. Muda status para PAGO
            abaTrans.getRange(i + 1, m['Status'] + 1).setValue('Pago');
            
            // 2. Define a Data que foi pago (Hoje)
            abaTrans.getRange(i + 1, m['Data_Pagamento'] + 1).setValue(dataPag);
            
            // 3. VINCULA A CONTA BANCÁRIA (Isso faz o saldo do PicPay descer na leitura do dashboard)
            // Agora essa despesa "pertence" ao PicPay
            abaTrans.getRange(i + 1, m['Conta_Origem'] + 1).setValue(dados.conta);
            
            totalBaixado += rValor;
            countItens++;
        }
    }

    if (totalBaixado === 0) {
        return { sucesso: false, erro: "Nenhum item em aberto encontrado para esta fatura." };
    }

    // 2. ATUALIZAR SALDO DA CONTA BANCÁRIA (PicPay)
    // Como editamos as transações colocando o ID da conta nelas, 
    // precisamos subtrair esse valor do saldo atual da conta em BD_Contas
    
    const mConta = getColMap(abaContas);
    const dataContas = abaContas.getDataRange().getValues();
    
    for (let i = 1; i < dataContas.length; i++) {
        if (String(dataContas[i][mConta['ID_Conta']]) === String(dados.conta)) {
            let saldoAtual = parseFloat(dataContas[i][mConta['Saldo_Atual']] || 0);
            let novoSaldo = saldoAtual - totalBaixado; // Subtrai o valor total pago
            
            abaContas.getRange(i + 1, mConta['Saldo_Atual'] + 1).setValue(novoSaldo);
            break;
        }
    }

    // 3. ATUALIZAR LIMITE USADO DO CARTÃO
    // Subtrai do "Total_Usado" ou similar em BD_Cartoes
    
    const mCart = getColMap(abaCartoes);
    const dataCart = abaCartoes.getDataRange().getValues();
    
    for (let i = 1; i < dataCart.length; i++) {
        if (String(dataCart[i][mCart['ID_Cartao']]) === String(dados.cartaoId)) {
            // Se tiver coluna 'Total_Usado' ou 'Limite_Utilizado'
            // O código usa 'Total Usado' no dashboard, mas vamos tentar achar a coluna certa
            let colIndex = -1;
            if (mCart['Total_Usado'] !== undefined) colIndex = mCart['Total_Usado'];
            else if (mCart['Limite_Utilizado'] !== undefined) colIndex = mCart['Limite_Utilizado'];
            
            if (colIndex !== -1) {
                let usoAtual = parseFloat(dataCart[i][colIndex] || 0);
                let novoUso = usoAtual - totalBaixado;
                if (novoUso < 0) novoUso = 0; // Segurança
                
                abaCartoes.getRange(i + 1, colIndex + 1).setValue(novoUso);
            }
            break;
        }
    }

    return { 
        sucesso: true, 
        msg: `${countItens} itens baixados. R$ ${totalBaixado.toFixed(2)} debitados da conta.` 
    };

  } catch (e) {
    return { sucesso: false, erro: "Erro no backend: " + e.toString() };
  }
}
