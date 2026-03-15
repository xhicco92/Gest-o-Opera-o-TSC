// Configurações
let dadosBrutos = [];
let ultimaAtualizacao = null;
let cabecalhosOriginais = [];

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Inicializando dashboard...');
    
    // Criar a estrutura das áreas
    criarEstruturaAreas();
    
    // Adicionar botão de upload
    adicionarBotaoUpload();
    
    // Botão de atualização
    document.getElementById('btnAtualizar').addEventListener('click', () => {
        document.getElementById('fileInput').click();
    });
    
    // Seletor de período
    document.getElementById('periodoSelect').addEventListener('change', (e) => {
        console.log('Período alterado:', e.target.value);
        carregarDadosExemplo();
    });
    
    // Carregar dados de exemplo inicialmente
    carregarDadosExemplo();
});

// Função para criar a estrutura das áreas dinamicamente
function criarEstruturaAreas() {
    const areasGrid = document.getElementById('areasGrid');
    
    const areas = [
        { id: 'mobileCliente', nome: '📱 Mobile Cliente', classe: 'mobile-cliente' },
        { id: 'mobileDG', nome: '📟 Mobile D&G', classe: 'mobile-dg' },
        { id: 'informatica', nome: '💻 Informática', classe: 'informatica' },
        { id: 'pequenos', nome: '🔌 Pequenos Domésticos', classe: 'pequenos-domesticos' },
        { id: 'som', nome: '🎵 Som e Imagem', classe: 'som-imagem' },
        { id: 'entretenimento', nome: '🎮 Entretenimento', classe: 'entretenimento' }
    ];
    
    const negociosPorArea = {
        'mobileCliente': [
            { id: 'Garantias', nome: 'Garantias', badge: 'garantias', label: 'G' },
            { id: 'ForaGarantia', nome: 'Fora de Garantia', badge: 'fora-garantia', label: 'FG' },
            { id: 'Extensao', nome: 'Extensão de Garantia', badge: 'extensao', label: 'EG' }
        ],
        'mobileDG': [
            { id: 'DG', nome: 'D&G', badge: 'dg', label: 'D&G' }
        ],
        'informatica': [
            { id: 'Garantias', nome: 'Garantias', badge: 'garantias', label: 'G' },
            { id: 'ForaGarantia', nome: 'Fora de Garantia', badge: 'fora-garantia', label: 'FG' },
            { id: 'Extensao', nome: 'Extensão de Garantia', badge: 'extensao', label: 'EG' }
        ],
        'pequenos': [
            { id: 'Garantias', nome: 'Garantias', badge: 'garantias', label: 'G' },
            { id: 'ForaGarantia', nome: 'Fora de Garantia', badge: 'fora-garantia', label: 'FG' },
            { id: 'Extensao', nome: 'Extensão de Garantia', badge: 'extensao', label: 'EG' }
        ],
        'som': [
            { id: 'Garantias', nome: 'Garantias', badge: 'garantias', label: 'G' },
            { id: 'ForaGarantia', nome: 'Fora de Garantia', badge: 'fora-garantia', label: 'FG' },
            { id: 'Extensao', nome: 'Extensão de Garantia', badge: 'extensao', label: 'EG' }
        ],
        'entretenimento': [
            { id: 'Garantias', nome: 'Garantias', badge: 'garantias', label: 'G' },
            { id: 'ForaGarantia', nome: 'Fora de Garantia', badge: 'fora-garantia', label: 'FG' },
            { id: 'Extensao', nome: 'Extensão de Garantia', badge: 'extensao', label: 'EG' }
        ]
    };
    
    const campos = [
        { id: 'Analises', label: 'Análises' },
        { id: 'Reparacao', label: 'Reparação' },
        { id: 'RepPendentePeca', label: 'Rep. Pendente Peça' },
        { id: 'RepNaoPendentePeca', label: 'Rep. Não Pendente Peça' },
        { id: 'TATAberto', label: 'TAT Aberto' },
        { id: 'Orcamento', label: 'Orçamento' },
        { id: 'AgAceitacao', label: 'Ag Aceitação Orçamento' },
        { id: 'Debitos', label: 'Débitos WSS' }
    ];
    
    areasGrid.innerHTML = '';
    
    areas.forEach(area => {
        const areaCard = document.createElement('div');
        areaCard.className = 'area-card';
        
        // Nome da área para o total
        const totalId = area.id === 'mobileDG' ? 'totalMobileDG' : 
                       `total${area.id.charAt(0).toUpperCase() + area.id.slice(1)}`;
        
        areaCard.innerHTML = `
            <div class="area-header ${area.classe}">
                <h2>${area.nome}</h2>
                <span class="area-total" id="${totalId}">0</span>
            </div>
            <div class="negocios-container" id="${area.id}Container"></div>
        `;
        
        areasGrid.appendChild(areaCard);
        
        const container = document.getElementById(`${area.id}Container`);
        const negocios = negociosPorArea[area.id];
        
        negocios.forEach(negocio => {
            const negocioRow = document.createElement('div');
            negocioRow.className = 'negocio-row';
            
            let kpisHTML = '';
            campos.forEach(campo => {
                const id = `${area.id}${negocio.id}${campo.id}`;
                kpisHTML += `
                    <div class="mini-kpi">
                        <span class="mini-label">${campo.label}</span>
                        <span class="mini-value" id="${id}">-</span>
                    </div>
                `;
            });
            
            negocioRow.innerHTML = `
                <div class="negocio-title">
                    <span class="badge ${negocio.badge}">${negocio.label}</span>
                    <span>${negocio.nome}</span>
                </div>
                <div class="negocio-kpis" style="grid-template-columns: repeat(8, 1fr);">
                    ${kpisHTML}
                </div>
            `;
            
            container.appendChild(negocioRow);
        });
    });
    
    console.log('✅ Estrutura de áreas criada');
}

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
    
    // Mapear índices
    const idxTipologia = encontrarIndice(cabecalhos, ['tipologia', 'Tipologia', 'area', 'Área']);
    const idxCheckpoint = encontrarIndice(cabecalhos, ['checkpoint_atual', 'checkpoint', 'Estado', 'status']);
    const idxPendentePeca = encontrarIndice(cabecalhos, ['pendente_peca', 'Pendente Peça', 'aguarda peça']);
    const idxDataCheckin = encontrarIndice(cabecalhos, ['data_checkin', 'checkin', 'Data Entrada']);
    const idxPolo = encontrarIndice(cabecalhos, ['polo', 'Polo', 'unidade']);
    const idxTipoGarantia = 8; // Coluna I
    
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
        
        // Calcular TAT
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
        let negocio = 'Outros';
        const tg = tipoGarantia.toLowerCase();
        
        if (tg.includes('pop') || tg.includes('stock de loja') || tg.includes('garantias')) {
            negocio = 'Garantias';
        } else if (tg.includes('fora de garantia') || tg.includes('instore mobility')) {
            negocio = 'Fora de Garantia';
        } else if (tg.includes('eg+1') || tg.includes('eg+3') || tg.includes('extensão')) {
            negocio = 'Extensão de Garantia';
        } else if (tg.includes('seguro d&g') || tg.includes('d&g')) {
            negocio = 'D&G';
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
            area: areaNorm,
            negocio: negocio,
            checkpoint: checkpoint.toLowerCase(),
            pendente_peca: isPendentePeca,
            tat: tat
        });
    }
    
    console.log(`✅ Processados ${dadosBrutos.length} registos`);
    
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

// Função para calcular as métricas
function calcularMetricas() {
    const estadosAnalise = ['análise técnica', 'analise tecnica', 'análise', 'analise'];
    const estadosReparacao = ['intervenção técnica', 'intervencao tecnica', 'reparação', 'reparacao'];
    const estadosTAT = ['análise técnica', 'analise tecnica', 'intervenção técnica', 'intervencao tecnica', 
                       'orçamento', 'orcamento', 'aguarda aceitação orçamento', 'aguarda aceitacao orcamento',
                       'nível 3', 'nivel 3', 'pré-análise', 'pre-analise', 'controlo de qualidade'];
    const estadosOrcamento = ['orçamento', 'orcamento'];
    const estadosAguarda = ['aguarda aceitação orçamento', 'aguarda aceitacao orcamento'];
    const estadosDebito = ['debit', 'débito', 'debito'];

    const metricas = {};
    
    // Inicializar estrutura
    const areas = ['Mobile Cliente', 'Mobile D&G', 'Informática', 'Pequenos Domésticos', 'Som e Imagem', 'Entretenimento'];
    const negocios = ['Garantias', 'Fora de Garantia', 'Extensão de Garantia', 'D&G'];
    
    areas.forEach(area => {
        metricas[area] = {};
        negocios.forEach(negocio => {
            metricas[area][negocio] = {
                analises: 0,
                reparacao: 0,
                repPendentePeca: 0,
                repNaoPendentePeca: 0,
                somaTat: 0,
                countTat: 0,
                tatAberto: 0,
                orcamento: 0,
                agAceitacao: 0,
                debitos: 0
            };
        });
    });

    // Processar registos
    dadosBrutos.forEach(item => {
        const area = item.area;
        const negocio = item.negocio;
        
        if (!metricas[area] || !metricas[area][negocio]) return;
        
        const stats = metricas[area][negocio];
        const checkpoint = item.checkpoint;
        
        // Análises
        if (estadosAnalise.some(estado => checkpoint.includes(estado))) {
            stats.analises++;
        }
        
        // Reparação
        if (estadosReparacao.some(estado => checkpoint.includes(estado))) {
            stats.reparacao++;
            if (item.pendente_peca) {
                stats.repPendentePeca++;
            } else {
                stats.repNaoPendentePeca++;
            }
        }
        
        // TAT
        if (estadosTAT.some(estado => checkpoint.includes(estado))) {
            stats.somaTat += item.tat;
            stats.countTat++;
        }
        
        // Orçamento
        if (estadosOrcamento.some(estado => checkpoint.includes(estado))) {
            stats.orcamento++;
        }
        
        // Aguarda
        if (estadosAguarda.some(estado => checkpoint.includes(estado))) {
            stats.agAceitacao++;
        }
        
        // Débitos
        if (estadosDebito.some(estado => checkpoint.includes(estado))) {
            stats.debitos++;
        }
    });

    // Calcular médias
    areas.forEach(area => {
        negocios.forEach(negocio => {
            const stats = metricas[area][negocio];
            if (stats.countTat > 0) {
                stats.tatAberto = Math.round((stats.somaTat / stats.countTat) * 10) / 10;
            }
        });
    });

    return metricas;
}

// Função para mostrar métricas
function calcularEMostrarMetricas() {
    const metricas = calcularMetricas();
    console.log('📊 Métricas calculadas:', metricas);
    
    // Mapeamento de nomes para IDs
    const mapaIds = {
        'Mobile Cliente': 'mobileCliente',
        'Mobile D&G': 'mobileDG',
        'Informática': 'informatica',
        'Pequenos Domésticos': 'pequenos',
        'Som e Imagem': 'som',
        'Entretenimento': 'entretenimento',
        'Garantias': 'Garantias',
        'Fora de Garantia': 'ForaGarantia',
        'Extensão de Garantia': 'Extensao',
        'D&G': 'DG'
    };
    
    // Atualizar cada campo
    Object.keys(metricas).forEach(area => {
        let totalArea = 0;
        
        Object.keys(metricas[area]).forEach(negocio => {
            const stats = metricas[area][negocio];
            if (!stats) return;
            
            totalArea += stats.analises || 0;
            
            const prefix = mapaIds[area];
            const negocioId = mapaIds[negocio];
            
            if (!prefix || !negocioId) return;
            
            const campos = [
                { id: `${prefix}${negocioId}Analises`, valor: stats.analises },
                { id: `${prefix}${negocioId}Reparacao`, valor: stats.reparacao },
                { id: `${prefix}${negocioId}RepPendentePeca`, valor: stats.repPendentePeca },
                { id: `${prefix}${negocioId}RepNaoPendentePeca`, valor: stats.repNaoPendentePeca },
                { id: `${prefix}${negocioId}TATAberto`, valor: stats.tatAberto ? stats.tatAberto.toFixed(1) : '0.0' },
                { id: `${prefix}${negocioId}Orcamento`, valor: stats.orcamento },
                { id: `${prefix}${negocioId}AgAceitacao`, valor: stats.agAceitacao },
                { id: `${prefix}${negocioId}Debitos`, valor: stats.debitos }
            ];
            
            campos.forEach(campo => {
                const element = document.getElementById(campo.id);
                if (element) {
                    element.textContent = campo.valor;
                }
            });
        });
        
        // Atualizar total da área
        const totalId = area === 'Mobile D&G' ? 'totalMobileDG' : 
                       `total${area.replace(' ', '')}`;
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
    
    Object.values(metricas).forEach(area => {
        Object.values(area).forEach(neg => {
            totalAnalises += neg.analises || 0;
            somaTatGlobal += neg.somaTat || 0;
            countTatGlobal += neg.countTat || 0;
            totalOrcamentos += neg.orcamento || 0;
            totalAguarda += neg.agAceitacao || 0;
        });
    });
    
    document.getElementById('totalEntradas').textContent = totalAnalises;
    document.getElementById('tatMedio').textContent = countTatGlobal > 0 ? (somaTatGlobal / countTatGlobal).toFixed(1) : '0';
    document.getElementById('taxaSucessoGlobal').textContent = totalAnalises > 0 ? Math.round((totalAnalises - totalAguarda) / totalAnalises * 100) : '0';
    document.getElementById('nssMedio').textContent = totalOrcamentos;
    document.getElementById('produtividadeGlobal').textContent = totalAguarda;
    
    // Rodapé
    document.getElementById('totalReparacoes').textContent = dadosBrutos.length;
    
    const estadosAbertos = ['análise técnica', 'intervenção técnica', 'orçamento', 'aguarda aceitação orçamento'];
    const emAberto = dadosBrutos.filter(item => 
        estadosAbertos.some(estado => item.checkpoint.includes(estado))
    ).length;
    document.getElementById('emAndamento').textContent = emAberto;
    
    document.getElementById('concluidasHoje').textContent = '-'; // Simplificado
    
    document.getElementById('ultimaAtualizacao').textContent = ultimaAtualizacao ? 
        ultimaAtualizacao.toLocaleString('pt-PT') : '-';
    document.getElementById('dataReferencia').textContent = `📅 ${new Date().toLocaleDateString('pt-PT')}`;
    
    console.log('✅ Dashboard atualizado!');
}

// Dados de exemplo
function carregarDadosExemplo() {
    dadosBrutos = [];
    const areas = ['Mobile Cliente', 'Mobile D&G', 'Informática', 'Pequenos Domésticos', 'Som e Imagem', 'Entretenimento'];
    const negociosLista = ['Garantias', 'Fora de Garantia', 'Extensão de Garantia', 'D&G'];
    const checkpoints = ['análise técnica', 'intervenção técnica', 'orçamento', 'aguarda aceitação orçamento', 'nível 3', 'pré-análise', 'controlo de qualidade', 'debit'];
    
    for (let i = 1; i <= 200; i++) {
        const area = areas[Math.floor(Math.random() * areas.length)];
        const negocio = negociosLista[Math.floor(Math.random() * negociosLista.length)];
        const checkpoint = checkpoints[Math.floor(Math.random() * checkpoints.length)];
        
        // Ajustar área para Mobile D&G quando negócio é D&G
        let areaFinal = area;
        if (negocio === 'D&G') {
            areaFinal = 'Mobile D&G';
        }
        
        dadosBrutos.push({
            area: areaFinal,
            negocio: negocio,
            checkpoint: checkpoint,
            pendente_peca: Math.random() > 0.7,
            tat: Math.random() * 15
        });
    }
    
    ultimaAtualizacao = new Date();
    calcularEMostrarMetricas();
    console.log('✅ Dados de exemplo carregados');
}
