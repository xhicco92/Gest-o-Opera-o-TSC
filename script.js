// Configurações
let dadosBrutos = [];
let ultimaAtualizacao = null;
let cabecalhosOriginais = [];
let dadosCarregados = false;

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Inicializando dashboard...');
    
    criarEstruturaAreas();
    adicionarBotaoUpload();
    
    document.getElementById('btnAtualizar').addEventListener('click', () => {
        document.getElementById('fileInput').click();
    });
    
    document.getElementById('periodoSelect').addEventListener('change', (e) => {
        console.log('Período alterado:', e.target.value);
    });
    
    ultimaAtualizacao = new Date();
    document.getElementById('ultimaAtualizacao').textContent = ultimaAtualizacao.toLocaleString('pt-PT');
    document.getElementById('dataReferencia').textContent = `📅 ${new Date().toLocaleDateString('pt-PT')}`;
    
    console.log('✅ Dashboard pronto - aguardando ficheiro Excel');
});

// Função para criar a estrutura das áreas
function criarEstruturaAreas() {
    const areasGrid = document.getElementById('areasGrid');
    if (!areasGrid) {
        console.error('❌ Elemento areasGrid não encontrado!');
        return;
    }
    
    console.log('🔄 Criando estrutura das áreas...');
    
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
                        <span class="mini-value" id="${id}">0</span>
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
    
    console.log('✅ Estrutura de áreas criada (tudo a 0)');
}

// Adicionar botão de upload
function adicionarBotaoUpload() {
    const headerControls = document.querySelector('.header-controls');
    if (!headerControls) return;
    
    const existingInput = document.getElementById('fileInput');
    if (existingInput) existingInput.remove();
    
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.id = 'fileInput';
    fileInput.accept = '.xlsx, .xls, .csv';
    fileInput.style.display = 'none';
    fileInput.addEventListener('change', handleFileUpload);
    
    headerControls.appendChild(fileInput);
    console.log('✅ Botão de upload adicionado');
}

// Processar upload
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    console.log('📂 Ficheiro selecionado:', file.name);
    dadosCarregados = true;
    
    try {
        document.getElementById('btnAtualizar').innerHTML = '<span class="material-icons">refresh</span> A processar...';
        
        const data = await lerFicheiroExcel(file);
        const sucesso = processarDadosExcel(data);
        
        event.target.value = '';
        
        if (sucesso) {
            ultimaAtualizacao = new Date();
            document.getElementById('ultimaAtualizacao').textContent = ultimaAtualizacao.toLocaleString('pt-PT');
            console.log('✅ Dados reais processados com sucesso!');
        } else {
            console.log('⚠️ Falha ao processar dados reais');
            dadosCarregados = false;
        }
        
    } catch (erro) {
        console.error('❌ Erro ao processar ficheiro:', erro);
        alert('Erro ao processar o ficheiro. Certifique-se que é um Excel válido.');
        dadosCarregados = false;
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

// Processar dados do Excel - VERSÃO CORRIGIDA (checkpoint = coluna 14)
function processarDadosExcel(dados) {
    if (!dados || dados.length < 2) {
        console.log('⚠️ Ficheiro sem dados suficientes');
        return false;
    }
    
    const cabecalhos = dados[0];
    cabecalhosOriginais = cabecalhos;
    console.log('📊 Cabeçalhos encontrados:', cabecalhos);
    
    // Índices fixos
    const idxTipologia = 33; // Coluna 34
    const idxTipoGarantia = 8; // Coluna 9
    const idxCheckpoint = 13; // Coluna 14 do Excel (ÍNDICE CORRETO!)
    
    // Encontrar outros índices
    const idxPendentePeca = encontrarIndice(cabecalhos, ['pendente_peca', 'Pendente Peça']);
    const idxDataCheckin = encontrarIndice(cabecalhos, ['data_checkin', 'checkin']);
    
    console.log('📍 Mapeamento:', {
        tipologia: idxTipologia,
        tipo_garantia: idxTipoGarantia,
        checkpoint: idxCheckpoint,
        pendente_peca: idxPendentePeca,
        data_checkin: idxDataCheckin
    });
    
    // Processar dados
    const novosDados = [];
    let linhasIgnoradas = 0;
    
    for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];
        if (!linha || linha.length === 0) continue;
        if (linha.every(cell => !cell || cell === '')) continue;
        
        // Extrair valores
        const tipologia = linha[idxTipologia] ? String(linha[idxTipologia]) : '';
        const tipoGarantia = linha[idxTipoGarantia] ? String(linha[idxTipoGarantia]) : '';
        const checkpointRaw = linha[idxCheckpoint] ? String(linha[idxCheckpoint]) : '';
        const pendentePeca = idxPendentePeca !== -1 ? String(linha[idxPendentePeca] || '').toLowerCase() : '';
        const dataCheckin = idxDataCheckin !== -1 ? linha[idxDataCheckin] : null;
        
        // Converter checkpoint para minúsculas e remover espaços extras
        const checkpoint = checkpointRaw.toLowerCase().trim();
        
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
            } catch (e) {
                linhasIgnoradas++;
            }
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
        }
        
        // Determinar a área
        let areaNorm = null;
        const tipologiaUpper = tipologia.toUpperCase();
        
        // Mobile
        if (tipologiaUpper.includes('MOBILE') || tipologiaUpper.includes('MÓVEL')) {
            if (tg.includes('seguro d&g')) {
                areaNorm = 'Mobile D&G';
                negocio = 'D&G';
            } else if (!tg.includes('flex premium') && !tg.includes('seguro d&g')) {
                areaNorm = 'Mobile Cliente';
            }
        }
        // Informática
        else if (tipologiaUpper.includes('INFORMATICA') || tipologiaUpper.includes('PC') || 
                 tipologiaUpper.includes('NOTEBOOK') || tipologiaUpper.includes('COMPUTADOR')) {
            areaNorm = 'Informática';
        }
        // Pequenos Domésticos
        else if (tipologiaUpper.includes('DOMESTICO') || tipologiaUpper.includes('PDA') || 
                 tipologiaUpper.includes('ELECTRO') || tipologiaUpper.includes('PEQUENO')) {
            areaNorm = 'Pequenos Domésticos';
        }
        // Som e Imagem
        else if (tipologiaUpper.includes('SOM') || tipologiaUpper.includes('IMAGEM') || 
                 tipologiaUpper.includes('TV') || tipologiaUpper.includes('AUDIO') ||
                 tipologiaUpper.includes('VIDEO')) {
            areaNorm = 'Som e Imagem';
        }
        // Entretenimento
        else if (tipologiaUpper.includes('ENTRETENIMENTO') || tipologiaUpper.includes('GAMING') || 
                 tipologiaUpper.includes('CONSOLA') || tipologiaUpper.includes('PLAYSTATION') ||
                 tipologiaUpper.includes('XBOX') || tipologiaUpper.includes('NINTENDO')) {
            areaNorm = 'Entretenimento';
        }
        
        // Só adicionar se a área foi identificada
        if (areaNorm) {
            novosDados.push({
                area: areaNorm,
                negocio: negocio,
                checkpoint: checkpoint,
                pendente_peca: isPendentePeca,
                tat: tat
            });
        }
    }
    
    if (novosDados.length === 0) {
        console.log('⚠️ Nenhum registo válido encontrado no ficheiro');
        return false;
    }
    
    // Substituir dados
    dadosBrutos = novosDados;
    
    console.log(`✅ Processados ${dadosBrutos.length} registos do ficheiro`);
    
    // Mostrar estatísticas detalhadas
    const contagemAreas = {};
    const contagemCheckpoints = {};
    dadosBrutos.forEach(item => {
        contagemAreas[item.area] = (contagemAreas[item.area] || 0) + 1;
        contagemCheckpoints[item.checkpoint] = (contagemCheckpoints[item.checkpoint] || 0) + 1;
    });
    
    console.log('📊 Distribuição por área:', contagemAreas);
    console.log('📌 Checkpoints encontrados:', contagemCheckpoints);
    
    // Atualizar dashboard
    calcularEMostrarMetricas();
    
    return true;
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
    // Estados que podem aparecer nos checkpoints
    const estadosAnalise = ['pré-análise', 'pre-analise', 'pré'];
    const estadosReparacao = ['intervenção técnica', 'intervencao tecnica', 'intervenção', 'reparação'];
    const estadosOrcamento = ['orçamento', 'orcamento'];
    const estadosAguarda = ['aguarda aceitação orçamento', 'aguarda aceitacao orcamento', 'aguarda'];
    const estadosDebito = ['debit', 'débito', 'debito'];

    // Inicializar tudo a zero
    const metricas = {
        'Mobile Cliente': {
            'Garantias': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 },
            'Fora de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 },
            'Extensão de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 }
        },
        'Mobile D&G': {
            'D&G': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 }
        },
        'Informática': {
            'Garantias': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 },
            'Fora de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 },
            'Extensão de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 }
        },
        'Pequenos Domésticos': {
            'Garantias': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 },
            'Fora de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 },
            'Extensão de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 }
        },
        'Som e Imagem': {
            'Garantias': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 },
            'Fora de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 },
            'Extensão de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 }
        },
        'Entretenimento': {
            'Garantias': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 },
            'Fora de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 },
            'Extensão de Garantia': { analises:0, reparacao:0, repPendentePeca:0, repNaoPendentePeca:0, somaTat:0, countTat:0, tatAberto:0, orcamento:0, agAceitacao:0, debitos:0 }
        }
    };

    // Processar cada registo
    dadosBrutos.forEach(item => {
        const area = item.area;
        const negocio = item.negocio;
        
        if (!metricas[area] || !metricas[area][negocio]) return;
        
        const stats = metricas[area][negocio];
        const checkpoint = String(item.checkpoint).toLowerCase();
        
        // Análises
        if (estadosAnalise.some(estado => checkpoint.includes(estado))) {
            stats.analises++;
        }
        
        // Reparação
        if (estadosReparacao.some(estado => checkpoint.includes(estado))) {
            stats.reparacao++;
            if (item.pendente_peca) stats.repPendentePeca++;
            else stats.repNaoPendentePeca++;
        }
        
        // TAT
        stats.somaTat += item.tat || 0;
        stats.countTat++;
        
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

// Função para mostrar métricas
function calcularEMostrarMetricas() {
    console.log('🔄 Atualizando dashboard...');
    
    const metricas = calcularMetricas();
    
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
    
    const combinacoesValidas = [
        { area: 'Mobile Cliente', negocios: ['Garantias', 'Fora de Garantia', 'Extensão de Garantia'] },
        { area: 'Mobile D&G', negocios: ['D&G'] },
        { area: 'Informática', negocios: ['Garantias', 'Fora de Garantia', 'Extensão de Garantia'] },
        { area: 'Pequenos Domésticos', negocios: ['Garantias', 'Fora de Garantia', 'Extensão de Garantia'] },
        { area: 'Som e Imagem', negocios: ['Garantias', 'Fora de Garantia', 'Extensão de Garantia'] },
        { area: 'Entretenimento', negocios: ['Garantias', 'Fora de Garantia', 'Extensão de Garantia'] }
    ];
    
    let elementosEncontrados = 0;
    let elementosTotal = 0;
    
    combinacoesValidas.forEach(combo => {
        const area = combo.area;
        const areaMetricas = metricas[area];
        if (!areaMetricas) return;
        
        let totalArea = 0;
        
        combo.negocios.forEach(negocio => {
            const stats = areaMetricas[negocio];
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
            
            elementosTotal += campos.length;
            
            campos.forEach(campo => {
                const element = document.getElementById(campo.id);
                if (element) {
                    element.textContent = campo.valor;
                    elementosEncontrados++;
                }
            });
        });
        
        const totalId = area === 'Mobile D&G' ? 'totalMobileDG' : `total${area.replace(' ', '').replace('&', '')}`;
        const totalElement = document.getElementById(totalId);
        if (totalElement) totalElement.textContent = totalArea;
    });
    
    console.log(`📊 Elementos atualizados: ${elementosEncontrados}/${elementosTotal}`);
    
    // KPIs globais
    let totalAnalises = 0, somaTatGlobal = 0, countTatGlobal = 0, totalOrcamentos = 0, totalAguarda = 0, totalDebitos = 0;
    
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
    
    // Rodapé
    document.getElementById('totalReparacoes').textContent = dadosBrutos.length;
    
    const estadosAbertos = ['análise', 'intervenção', 'orçamento', 'aguarda'];
    const emAberto = dadosBrutos.filter(item => 
        estadosAbertos.some(s => item.checkpoint.includes(s))
    ).length;
    document.getElementById('emAndamento').textContent = emAberto;
    
    document.getElementById('concluidasHoje').textContent = '-';
    document.getElementById('ultimaAtualizacao').textContent = ultimaAtualizacao ? ultimaAtualizacao.toLocaleString('pt-PT') : '-';
    document.getElementById('dataReferencia').textContent = `📅 ${new Date().toLocaleDateString('pt-PT')}`;
    
    if (dadosBrutos.length > 0) {
        console.log('✅ Dashboard atualizado com DADOS REAIS');
    } else {
        console.log('✅ Dashboard atualizado (sem dados - tudo a 0)');
    }
}
