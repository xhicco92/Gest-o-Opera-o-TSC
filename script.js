// Configurações
let dadosBrutos = [];
let ultimaAtualizacao = null;
let cabecalhosOriginais = [];

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Inicializando dashboard com cálculos...');
    
    adicionarBotaoUpload();
    
    document.getElementById('btnAtualizar').addEventListener('click', () => {
        document.getElementById('fileInput').click();
    });
    
    document.getElementById('periodoSelect').addEventListener('change', (e) => {
        console.log('Período alterado:', e.target.value);
        carregarDadosExemplo();
    });
    
    carregarDadosExemplo();
});

// Adicionar botão de upload
function adicionarBotaoUpload() {
    const headerControls = document.querySelector('.header-controls');
    
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.id = 'fileInput';
    fileInput.accept = '.xlsx, .xls, .csv';
    fileInput.style.display = 'none';
    fileInput.addEventListener('change', handleFileUpload);
    
    headerControls.appendChild(fileInput);
}

// Processar upload
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    console.log('📂 Ficheiro selecionado:', file.name);
    
    try {
        document.getElementById('btnAtualizar').innerHTML = '<span class="material-icons">refresh</span> A processar...';
        
        const data = await lerFicheiroExcel(file);
        processarDadosExcel(data);
        
        event.target.value = '';
        
        ultimaAtualizacao = new Date();
        document.getElementById('ultimaAtualizacao').textContent = ultimaAtualizacao.toLocaleString('pt-PT');
        
        console.log('✅ Dados processados com sucesso!');
        
    } catch (erro) {
        console.error('❌ Erro ao processar ficheiro:', erro);
        alert('Erro ao processar o ficheiro. Certifique-se que é um Excel válido.');
    } finally {
        document.getElementById('btnAtualizar').innerHTML = '<span class="material-icons">refresh</span> Carregar Excel';
    }
}

// Ler ficheiro Excel
async function lerFicheiroExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                if (!workbook.SheetNames.includes('Dados')) {
                    reject(new Error('Folha "Dados" não encontrada'));
                    return;
                }
                
                const worksheet = workbook.Sheets['Dados'];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1,
                    defval: '',
                    range: 0
                });
                
                resolve(jsonData);
                
            } catch (erro) {
                reject(erro);
            }
        };
        
        reader.onerror = (erro) => reject(erro);
        reader.readAsArrayBuffer(file);
    });
}

// Processar dados do Excel
function processarDadosExcel(dados) {
    if (!dados || dados.length < 2) {
        console.log('⚠️ Ficheiro sem dados suficientes');
        carregarDadosExemplo();
        return;
    }
    
    const cabecalhos = dados[0];
    cabecalhosOriginais = cabecalhos;
    console.log('📊 Cabeçalhos encontrados:', cabecalhos);
    
    // Mapear índices (coluna I = índice 8, zero-based)
    const idxTipologia = encontrarIndice(cabecalhos, ['tipologia', 'Tipologia', 'area', 'Área']);
    const idxCheckpoint = encontrarIndice(cabecalhos, ['checkpoint_atual', 'checkpoint', 'Estado', 'status']);
    const idxPendentePeca = encontrarIndice(cabecalhos, ['pendente_peca', 'Pendente Peça', 'aguarda peça']);
    const idxDataCheckin = encontrarIndice(cabecalhos, ['data_checkin', 'checkin', 'Data Entrada']);
    const idxPolo = encontrarIndice(cabecalhos, ['polo', 'Polo', 'unidade']);
    const idxTipoGarantia = 8; // Coluna I (índice 8, zero-based)
    
    console.log('📍 Mapeamento:', {
        tipologia: idxTipologia,
        checkpoint: idxCheckpoint,
        pendente_peca: idxPendentePeca,
        data_checkin: idxDataCheckin,
        polo: idxPolo,
        tipo_garantia: idxTipoGarantia
    });
    
    // Processar dados
    dadosBrutos = [];
    
    for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];
        if (!linha || linha.length === 0) continue;
        if (linha.every(cell => !cell || cell === '')) continue;
        
        // Filtrar TSC South
        const polo = idxPolo !== -1 ? String(linha[idxPolo] || '').toLowerCase() : '';
        if (!polo.includes('tsc south') && !polo.includes('south')) {
            continue;
        }
        
        // Extrair valores
        const tipologia = idxTipologia !== -1 ? String(linha[idxTipologia] || '') : '';
        const checkpoint = idxCheckpoint !== -1 ? String(linha[idxCheckpoint] || '') : '';
        const pendentePeca = idxPendentePeca !== -1 ? String(linha[idxPendentePeca] || '').toLowerCase() : '';
        const dataCheckin = idxDataCheckin !== -1 ? linha[idxDataCheckin] : null;
        const tipoGarantia = linha[idxTipoGarantia] ? String(linha[idxTipoGarantia]) : '';
        
        // Calcular TAT (dias desde data_checkin)
        let tat = 0;
        if (dataCheckin) {
            try {
                let dataEntrada;
                if (typeof dataCheckin === 'number') {
                    dataEntrada = new Date((dataCheckin - 25569) * 86400 * 1000);
                } else {
                    dataEntrada = new Date(dataCheckin);
                }
                if (!isNaN(dataEntrada)) {
                    const hoje = new Date();
                    tat = Math.abs(hoje - dataEntrada) / (1000 * 60 * 60 * 24);
                }
            } catch (e) {}
        }
        
        // Determinar se pendente peça é Sim/Não
        const isPendentePeca = pendentePeca.includes('sim') || pendentePeca === 's' || pendentePeca === '1';
        
        // Mapear tipo_garantia para negócio
        let negocio = '';
        const tg = tipoGarantia.toLowerCase();
        
        if (tg.includes('pop') || tg.includes('stock de loja') || tg.includes('garantias')) {
            negocio = 'Garantias';
        } else if (tg.includes('fora de garantia') || tg.includes('instore mobility')) {
            negocio = 'Fora de Garantia';
        } else if (tg.includes('eg+1') || tg.includes('eg+3') || tg.includes('extensão')) {
            negocio = 'Extensão de Garantia';
        } else if (tg.includes('seguro d&g') || tg.includes('d&g')) {
            negocio = 'D&G';
        } else {
            negocio = 'Outros';
        }
        
        // Normalizar tipologia para as áreas do dashboard
        let areaNorm = 'Outros';
        const tipologiaUpper = tipologia.toUpperCase();
        
        if (tipologiaUpper.includes('MOBILE') || tipologiaUpper.includes('TELEMOVEL')) {
            if (tipologiaUpper.includes('D&G') || tipologiaUpper.includes('DG') || negocio === 'D&G') {
                areaNorm = 'Mobile D&G';
            } else {
                areaNorm = 'Mobile Cliente';
            }
        } else if (tipologiaUpper.includes('INFORMATICA') || tipologiaUpper.includes('PC') || tipologiaUpper.includes('NOTEBOOK')) {
            areaNorm = 'Informática';
        } else if (tipologiaUpper.includes('DOMESTICO') || tipologiaUpper.includes('PDA') || tipologiaUpper.includes('ELECTRO')) {
            areaNorm = 'Pequenos Domésticos';
        } else if (tipologiaUpper.includes('SOM') || tipologiaUpper.includes('IMAGEM') || tipologiaUpper.includes('TV') || tipologiaUpper.includes('AUDIO')) {
            areaNorm = 'Som e Imagem';
        } else if (tipologiaUpper.includes('ENTRETENIMENTO') || tipologiaUpper.includes('GAMING') || tipologiaUpper.includes('CONSOLA')) {
            areaNorm = 'Entretenimento';
        }
        
        // Ajustar área para Mobile D&G quando negócio é D&G
        if (negocio === 'D&G') {
            areaNorm = 'Mobile D&G';
        }
        
        dadosBrutos.push({
            id: dadosBrutos.length + 1,
            area: areaNorm,
            negocio: negocio,
            tipologia_original: tipologia,
            tipo_garantia_original: tipoGarantia,
            checkpoint: checkpoint,
            pendente_peca: isPendentePeca,
            data_checkin: dataCheckin,
            tat: tat,
            polo: polo
        });
    }
    
    console.log(`✅ Processados ${dadosBrutos.length} registos (TSC South apenas)`);
    console.log('📊 Distribuição por negócio:', {
        garantias: dadosBrutos.filter(d => d.negocio === 'Garantias').length,
        foraGarantia: dadosBrutos.filter(d => d.negocio === 'Fora de Garantia').length,
        extensao: dadosBrutos.filter(d => d.negocio === 'Extensão de Garantia').length,
        dg: dadosBrutos.filter(d => d.negocio === 'D&G').length,
        outros: dadosBrutos.filter(d => d.negocio === 'Outros').length
    });
    
    if (dadosBrutos.length > 0) {
        ultimaAtualizacao = new Date();
        calcularEMostrarMetricas();
    } else {
        carregarDadosExemplo();
    }
}

// Função auxiliar para encontrar índice
function encontrarIndice(cabecalhos, possiveisNomes) {
    for (let i = 0; i < cabecalhos.length; i++) {
        const cab = String(cabecalhos[i] || '').toLowerCase().trim();
        for (let nome of possiveisNomes) {
            if (cab.includes(nome.toLowerCase())) {
                console.log(`✅ Coluna "${cabecalhos[i]}" corresponde a "${nome}" (índice ${i})`);
                return i;
            }
        }
    }
    return -1;
}

// Função para calcular as métricas por área e negócio
function calcularMetricas() {
    // Mapeamento de checkpoints para estados
    const estadosAnalise = ['análise técnica', 'analise tecnica', 'análise', 'analise'];
    const estadosReparacao = ['intervenção técnica', 'intervencao tecnica', 'reparação', 'reparacao'];
    const estadosTAT = ['análise técnica', 'analise tecnica', 'intervenção técnica', 'intervencao tecnica', 
                       'orçamento', 'orcamento', 'aguarda aceitação orçamento', 'aguarda aceitacao orcamento',
                       'nível 3', 'nivel 3', 'pré-análise', 'pre-analise', 'controlo de qualidade'];
    const estadosOrcamento = ['orçamento', 'orcamento'];
    const estadosAguarda = ['aguarda aceitação orçamento', 'aguarda aceitacao orcamento'];
    const estadosDebito = ['debit', 'débito', 'debito'];

    // Estrutura para armazenar métricas por área e negócio
    const metricas = {
        'Mobile Cliente': {
            'Garantias': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 },
            'Fora de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 },
            'Extensão de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 }
        },
        'Mobile D&G': {
            'D&G': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 }
        },
        'Informática': {
            'Garantias': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 },
            'Fora de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 },
            'Extensão de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 }
        },
        'Pequenos Domésticos': {
            'Garantias': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 },
            'Fora de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 },
            'Extensão de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 }
        },
        'Som e Imagem': {
            'Garantias': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 },
            'Fora de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 },
            'Extensão de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 }
        },
        'Entretenimento': {
            'Garantias': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 },
            'Fora de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 },
            'Extensão de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0, somaTat:0, countTat:0 }
        }
    };

    // Processar cada registo
    dadosBrutos.forEach(item => {
        const area = item.area;
        if (!metricas[area]) return;
        
        const negocio = item.negocio;
        if (!metricas[area][negocio]) return;
        
        const stats = metricas[area][negocio];
        const checkpoint = item.checkpoint.toLowerCase();
        
        // Análises
        if (estadosAnalise.some(estado => checkpoint.includes(estado))) {
            stats.analises++;
        }
        
        // Reparação
        if (estadosReparacao.some(estado => checkpoint.includes(estado))) {
            stats.reparacao++;
            
            // Pendente/Não Pendente Peça
            if (item.pendente_peca) {
                stats.repPendentePeca++;
            } else {
                stats.repNaoPendentePeca++;
            }
        }
        
        // TAT Aberto (soma para cálculo da média)
        if (estadosTAT.some(estado => checkpoint.includes(estado))) {
            stats.somaTat += item.tat;
            stats.countTat++;
        }
        
        // Orçamento
        if (estadosOrcamento.some(estado => checkpoint.includes(estado))) {
            stats.orcamento++;
        }
        
        // Aguarda Aceitação
        if (estadosAguarda.some(estado => checkpoint.includes(estado))) {
            stats.agAceitacao++;
        }
        
        // Débitos
        if (estadosDebito.some(estado => checkpoint.includes(estado))) {
            stats.debitos++;
        }
    });

    // Calcular médias do TAT
    Object.keys(metricas).forEach(area => {
        Object.keys(metricas[area]).forEach(negocio => {
            const stats = metricas[area][negocio];
            if (stats.countTat > 0) {
                stats.tatAberto = Math.round((stats.somaTat / stats.countTat) * 10) / 10;
            }
        });
    });

    return metricas;
}

// Função para mostrar as métricas no dashboard
function calcularEMostrarMetricas() {
    const metricas = calcularMetricas();
    console.log('📊 Métricas calculadas:', metricas);
    
    // Atualizar cada campo no HTML
    const areas = [
        { nome: 'Mobile Cliente', prefix: 'mobileCliente', negocios: ['Garantias', 'ForaGarantia', 'Extensao'] },
        { nome: 'Mobile D&G', prefix: 'mobileDG', negocios: ['DG'] },
        { nome: 'Informática', prefix: 'informatica', negocios: ['Garantias', 'ForaGarantia', 'Extensao'] },
        { nome: 'Pequenos Domésticos', prefix: 'pequenos', negocios: ['Garantias', 'ForaGarantia', 'Extensao'] },
        { nome: 'Som e Imagem', prefix: 'som', negocios: ['Garantias', 'ForaGarantia', 'Extensao'] },
        { nome: 'Entretenimento', prefix: 'entretenimento', negocios: ['Garantias', 'ForaGarantia', 'Extensao'] }
    ];
    
    // Mapeamento de nomes de negócio para chaves nas métricas
    const mapaNegocios = {
        'Garantias': 'Garantias',
        'ForaGarantia': 'Fora de Garantia',
        'Extensao': 'Extensão de Garantia',
        'DG': 'D&G'
    };
    
    areas.forEach(areaConfig => {
        const areaMetricas = metricas[areaConfig.nome];
        if (!areaMetricas) return;
        
        // Total da área (soma de todas as análises)
        let totalArea = 0;
        
        areaConfig.negocios.forEach(negocioKey => {
            const negocioNome = mapaNegocios[negocioKey];
            const stats = areaMetricas[negocioNome];
            if (!stats) return;
            
            totalArea += stats.analises || 0;
            
            // Determinar prefixo correto para os IDs
            let prefix = areaConfig.prefix;
            
            // Atualizar cada campo
            const campos = [
                { id: `${prefix}${negocioKey}Analises`, valor: stats.analises },
                { id: `${prefix}${negocioKey}Reparacao`, valor: stats.reparacao },
                { id: `${prefix}${negocioKey}RepPendentePeca`, valor: stats.repPendentePeca },
                { id: `${prefix}${negocioKey}RepNaoPendentePeca`, valor: stats.repNaoPendentePeca },
                { id: `${prefix}${negocioKey}TATAberto`, valor: stats.tatAberto ? stats.tatAberto.toFixed(1) : '0.0' },
                { id: `${prefix}${negocioKey}Orcamento`, valor: stats.orcamento },
                { id: `${prefix}${negocioKey}AgAceitacao`, valor: stats.agAceitacao },
                { id: `${prefix}${negocioKey}Debitos`, valor: stats.debitos }
            ];
            
            campos.forEach(campo => {
                const element = document.getElementById(campo.id);
                if (element) {
                    element.textContent = campo.valor;
                } else {
                    console.log(`Elemento não encontrado: ${campo.id}`);
                }
            });
        });
        
        // Atualizar total da área
        const totalId = `total${areaConfig.prefix === 'mobileDG' ? 'MobileDG' : 
            (areaConfig.prefix.charAt(0).toUpperCase() + areaConfig.prefix.slice(1))}`;
        const totalElement = document.getElementById(totalId);
        if (totalElement) {
            totalElement.textContent = totalArea;
        }
    });
    
    // Calcular KPIs globais
    let totalAnalises = 0;
    let somaTatGlobal = 0;
    let countTatGlobal = 0;
    let totalOrcamentos = 0;
    let totalAguarda = 0;
    let totalDebitos = 0;
    
    Object.values(metricas).forEach(area => {
        Object.values(area).forEach(neg => {
            totalAnalises += neg.analises || 0;
            somaTatGlobal += neg.somaTat || 0;
            countTatGlobal += neg.countTat || 0;
            totalOrcamentos += neg.orcamento || 0;
            totalAguarda += neg.agAceitacao || 0;
            totalDebitos += neg.debitos || 0;
        });
    });
    
    document.getElementById('totalEntradas').textContent = totalAnalises;
    document.getElementById('tatMedio').textContent = countTatGlobal > 0 ? (somaTatGlobal / countTatGlobal).toFixed(1) : '0';
    document.getElementById('taxaSucessoGlobal').textContent = totalAnalises > 0 ? Math.round((totalAnalises - totalAguarda) / totalAnalises * 100) : '0';
    document.getElementById('nssMedio').textContent = totalOrcamentos;
    document.getElementById('produtividadeGlobal').textContent = totalAguarda;
    
    // Atualizar rodapé
    document.getElementById('totalReparacoes').textContent = dadosBrutos.length;
    
    const estadosAbertos = ['análise técnica', 'intervenção técnica', 'orçamento', 'aguarda aceitação orçamento'];
    const emAberto = dadosBrutos.filter(item => 
        estadosAbertos.some(estado => item.checkpoint.toLowerCase().includes(estado))
    ).length;
    document.getElementById('emAndamento').textContent = emAberto;
    
    const hoje = new Date().toISOString().split('T')[0];
    const checkinsHoje = dadosBrutos.filter(item => {
        if (!item.data_checkin) return false;
        let dataStr;
        if (typeof item.data_checkin === 'number') {
            dataStr = new Date((item.data_checkin - 25569) * 86400 * 1000).toISOString().split('T')[0];
        } else {
            dataStr = new Date(item.data_checkin).toISOString().split('T')[0];
        }
        return dataStr === hoje;
    }).length;
    document.getElementById('concluidasHoje').textContent = checkinsHoje;
    
    document.getElementById('ultimaAtualizacao').textContent = ultimaAtualizacao ? 
        ultimaAtualizacao.toLocaleString('pt-PT') : '-';
    document.getElementById('dataReferencia').textContent = `📅 ${new Date().toLocaleDateString('pt-PT')}`;
    
    console.log('✅ Dashboard atualizado com os dados reais!');
}

// Dados de exemplo (fallback)
function carregarDadosExemplo() {
    dadosBrutos = [];
    const areas = ['Mobile Cliente', 'Mobile D&G', 'Informática', 'Pequenos Domésticos', 'Som e Imagem', 'Entretenimento'];
    const negociosLista = ['Garantias', 'Fora de Garantia', 'Extensão de Garantia', 'D&G'];
    const checkpoints = ['Análise Técnica', 'Intervenção Técnica', 'Orçamento', 'Aguarda Aceitação Orçamento', 'Nível 3', 'Pré-Análise', 'Controlo de Qualidade', 'Debit'];
    
    for (let i = 1; i <= 200; i++) {
        const area = areas[Math.floor(Math.random() * areas.length)];
        const negocio = negociosLista[Math.floor(Math.random() * negociosLista.length)];
        const checkpoint = checkpoints[Math.floor(Math.random() * checkpoints.length)];
        const data = new Date();
        data.setDate(data.getDate() - Math.floor(Math.random() * 30));
        
        // Ajustar área para Mobile D&G quando negócio é D&G
        let areaFinal = area;
        if (negocio === 'D&G') {
            areaFinal = 'Mobile D&G';
        }
        
        dadosBrutos.push({
            id: i,
            area: areaFinal,
            negocio: negocio,
            checkpoint: checkpoint,
            pendente_peca: Math.random() > 0.7,
            data_checkin: data,
            tat: Math.random() * 15,
            polo: 'TSC South'
        });
    }
    
    ultimaAtualizacao = new Date();
    calcularEMostrarMetricas();
    console.log('✅ Dados de exemplo carregados');
}
