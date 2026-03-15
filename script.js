// Configurações
let dadosBrutos = [];
let ultimaAtualizacao = null;
let cabecalhosOriginais = [];

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Inicializando dashboard com nova estrutura...');
    
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
                
                // Usar folha "Dados"
                if (!workbook.SheetNames.includes('Dados')) {
                    reject(new Error('Folha "Dados" não encontrada'));
                    return;
                }
                
                const worksheet = workbook.Sheets['Dados'];
                
                // Converter para JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1,
                    defval: '',
                    range: 0 // Começa na linha 1
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
    
    // Cabeçalhos na primeira linha
    const cabecalhos = dados[0];
    cabecalhosOriginais = cabecalhos;
    console.log('📊 Cabeçalhos encontrados:', cabecalhos);
    
    // Mapear índices
    const idxPolo = encontrarIndice(cabecalhos, ['polo', 'Polo', 'unidade']);
    const idxArea = encontrarIndice(cabecalhos, ['area', 'Área', 'departamento']);
    const idxGarantia = encontrarIndice(cabecalhos, ['tipo_garantia', 'garantia', 'Tipo Garantia']);
    const idxData = encontrarIndice(cabecalhos, ['data_checkin', 'checkin', 'Data Entrada']);
    const idxPendentePeca = encontrarIndice(cabecalhos, ['pendente_peca', 'Pendente Peça', 'aguarda peça']);
    const idxAguardaCotacao = encontrarIndice(cabecalhos, ['aguarda_cotacao_de_peca', 'aguarda cotação', 'cotação']);
    const idxTipoReparacao = encontrarIndice(cabecalhos, ['tipo_reparacao', 'Tipo Reparação']);
    const idxStatus = encontrarIndice(cabecalhos, ['checkpoint_atual', 'status', 'Estado']);
    
    console.log('📍 Mapeamento:', {
        polo: idxPolo,
        area: idxArea,
        garantia: idxGarantia,
        data: idxData,
        pendente_peca: idxPendentePeca,
        aguarda_cotacao: idxAguardaCotacao,
        tipo_reparacao: idxTipoReparacao,
        status: idxStatus
    });
    
    // Processar dados (linhas a partir da 2)
    dadosBrutos = [];
    
    for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];
        if (!linha || linha.length === 0) continue;
        if (linha.every(cell => !cell || cell === '')) continue;
        
        // Filtrar apenas TSC South
        const polo = idxPolo !== -1 ? String(linha[idxPolo] || '').toLowerCase() : '';
        if (!polo.includes('tsc south') && !polo.includes('south')) {
            continue;
        }
        
        // Extrair valores
        const area = idxArea !== -1 ? linha[idxArea] : '';
        const tipoGarantia = idxGarantia !== -1 ? linha[idxGarantia] : '';
        const dataCheckin = idxData !== -1 ? linha[idxData] : null;
        const pendentePeca = idxPendentePeca !== -1 ? String(linha[idxPendentePeca] || '').toLowerCase() : '';
        const aguardaCotacao = idxAguardaCotacao !== -1 ? String(linha[idxAguardaCotacao] || '').toLowerCase() : '';
        const tipoReparacao = idxTipoReparacao !== -1 ? String(linha[idxTipoReparacao] || '').toLowerCase() : '';
        const checkpoint = idxStatus !== -1 ? linha[idxStatus] : '';
        
        // Determinar status
        let status = 'pendente';
        if (checkpoint) {
            const cp = String(checkpoint).toUpperCase();
            if (cp.includes('FECHADO') || cp.includes('CONCLUIDO')) {
                status = 'concluido';
            } else if (cp.includes('REPARACAO') || cp.includes('ANALISE')) {
                status = 'andamento';
            }
        }
        
        // Calcular TAT
        let tempoReparo = 0;
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
                    tempoReparo = Math.abs(hoje - dataEntrada) / (1000 * 60 * 60 * 24);
                }
            } catch (e) {}
        }
        
        // Normalizar área (igual ao anterior)
        let areaNorm = 'Outros';
        const areaStr = String(area || '').toUpperCase();
        
        if (areaStr.includes('MOBILE') || areaStr.includes('TELEMOVEL')) {
            if (areaStr.includes('D&G') || areaStr.includes('DG')) {
                areaNorm = 'Mobile D&G';
            } else {
                areaNorm = 'Mobile Cliente';
            }
        } else if (areaStr.includes('INFORMATICA') || areaStr.includes('PC') || areaStr.includes('NOTEBOOK')) {
            areaNorm = 'Informática';
        } else if (areaStr.includes('DOMESTICO') || areaStr.includes('PDA') || areaStr.includes('ELECTRO')) {
            areaNorm = 'Pequenos Domésticos';
        } else if (areaStr.includes('SOM') || areaStr.includes('IMAGEM') || areaStr.includes('TV') || areaStr.includes('AUDIO')) {
            areaNorm = 'Som e Imagem';
        } else if (areaStr.includes('ENTRETENIMENTO') || areaStr.includes('GAMING') || areaStr.includes('CONSOLA')) {
            areaNorm = 'Entretenimento';
        }
        
        // Normalizar negócio
        let negocioNorm = 'Garantias';
        const tg = String(tipoGarantia || '').toUpperCase();
        if (tg.includes('FORA') || tg.includes('PAGO')) {
            negocioNorm = 'Fora de Garantia';
        } else if (tg.includes('EXTENSAO') || tg.includes('PROTECAO')) {
            negocioNorm = 'Extensão de Garantia';
        }
        
        if (areaNorm === 'Mobile D&G') {
            negocioNorm = 'D&G';
        }
        
        // Mapear flags
        const isPendentePeca = pendentePeca.includes('sim') || pendentePeca.includes('s') || pendentePeca === '1';
        const isAguardaCotacao = aguardaCotacao.includes('sim') || aguardaCotacao.includes('s') || aguardaCotacao === '1';
        const isOrcamento = tipoReparacao.includes('orcamento') || tipoReparacao.includes('orçamento');
        
        dadosBrutos.push({
            id: dadosBrutos.length + 1,
            area: areaNorm,
            negocio: negocioNorm,
            status: status,
            data_entrada: dataCheckin,
            tempo_reparo: Math.round(tempoReparo * 10) / 10,
            pendente_peca: isPendentePeca,
            aguarda_cotacao: isAguardaCotacao,
            is_orcamento: isOrcamento,
            tipo_garantia: tipoGarantia,
            checkpoint: checkpoint,
            polo: polo
        });
    }
    
    console.log(`✅ Processados ${dadosBrutos.length} registos (TSC South apenas)`);
    
    if (dadosBrutos.length > 0) {
        ultimaAtualizacao = new Date();
        atualizarInterface();
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
                return i;
            }
        }
    }
    return -1;
}

// Função para calcular todas as métricas
function calcularMetricas() {
    const metricas = {
        mobileCliente: { 
            global: 0,
            analises: { Garantias:0, ForaGarantia:0, Extensao:0 },
            reparacoes: { global:0, pendentesPeca:0, naoPendentesPeca:0 },
            tatAberto: 0,
            orcamentos: 0,
            aguardaAceitacao: 0,
            debitosWSS: 0,
            somaTat: 0,
            countTat: 0
        },
        mobileDG: { 
            global: 0,
            analises: { Garantias:0, ForaGarantia:0, Extensao:0 },
            reparacoes: { global:0, pendentesPeca:0, naoPendentesPeca:0 },
            tatAberto: 0,
            orcamentos: 0,
            aguardaAceitacao: 0,
            debitosWSS: 0,
            somaTat: 0,
            countTat: 0
        },
        informatica: { 
            global: 0,
            analises: { Garantias:0, ForaGarantia:0, Extensao:0 },
            reparacoes: { global:0, pendentesPeca:0, naoPendentesPeca:0 },
            tatAberto: 0,
            orcamentos: 0,
            aguardaAceitacao: 0,
            debitosWSS: 0,
            somaTat: 0,
            countTat: 0
        },
        pequenosDomesticos: { 
            global: 0,
            analises: { Garantias:0, ForaGarantia:0, Extensao:0 },
            reparacoes: { global:0, pendentesPeca:0, naoPendentesPeca:0 },
            tatAberto: 0,
            orcamentos: 0,
            aguardaAceitacao: 0,
            debitosWSS: 0,
            somaTat: 0,
            countTat: 0
        },
        somImagem: { 
            global: 0,
            analises: { Garantias:0, ForaGarantia:0, Extensao:0 },
            reparacoes: { global:0, pendentesPeca:0, naoPendentesPeca:0 },
            tatAberto: 0,
            orcamentos: 0,
            aguardaAceitacao: 0,
            debitosWSS: 0,
            somaTat: 0,
            countTat: 0
        },
        entretenimento: { 
            global: 0,
            analises: { Garantias:0, ForaGarantia:0, Extensao:0 },
            reparacoes: { global:0, pendentesPeca:0, naoPendentesPeca:0 },
            tatAberto: 0,
            orcamentos: 0,
            aguardaAceitacao: 0,
            debitosWSS: 0,
            somaTat: 0,
            countTat: 0
        },
        outros: { 
            global: 0,
            analises: { Garantias:0, ForaGarantia:0, Extensao:0 },
            reparacoes: { global:0, pendentesPeca:0, naoPendentesPeca:0 },
            tatAberto: 0,
            orcamentos: 0,
            aguardaAceitacao: 0,
            debitosWSS: 0,
            somaTat: 0,
            countTat: 0
        }
    };

    const mapaAreas = {
        'Mobile Cliente': 'mobileCliente',
        'Mobile D&G': 'mobileDG',
        'Informática': 'informatica',
        'Pequenos Domésticos': 'pequenosDomesticos',
        'Som e Imagem': 'somImagem',
        'Entretenimento': 'entretenimento',
        'Outros': 'outros'
    };

    // Processar cada registo
    dadosBrutos.forEach(item => {
        const areaKey = mapaAreas[item.area] || 'outros';
        const area = metricas[areaKey];
        
        // Global
        area.global++;
        
        // Análises por tipo de garantia
        if (item.negocio === 'Garantias') area.analises.Garantias++;
        else if (item.negocio === 'Fora de Garantia') area.analises.ForaGarantia++;
        else if (item.negocio === 'Extensão de Garantia') area.analises.Extensao++;
        else if (item.negocio === 'D&G') area.analises.Garantias++; // D&G conta como Garantias
        
        // Reparações
        if (item.status !== 'concluido') {
            area.reparacoes.global++;
            if (item.pendente_peca) {
                area.reparacoes.pendentesPeca++;
            } else {
                area.reparacoes.naoPendentesPeca++;
            }
        }
        
        // TAT Aberto
        if (item.tempo_reparo > 0) {
            area.somaTat += item.tempo_reparo;
            area.countTat++;
        }
        
        // Orçamentos
        if (item.is_orcamento) {
            area.orcamentos++;
        }
        
        // Aguarda Aceitação (pendente_peca + aguarda_cotacao)
        if (item.aguarda_cotacao) {
            area.aguardaAceitacao++;
        }
        
        // Débitos WSS (por definir - talvez todos os registos?)
        area.debitosWSS = area.global; // Temporário
    });

    // Calcular médias
    Object.keys(metricas).forEach(key => {
        const area = metricas[key];
        if (area.countTat > 0) {
            area.tatAberto = Math.round((area.somaTat / area.countTat) * 10) / 10;
        }
    });

    return metricas;
}

// Função para atualizar a interface com a nova estrutura
function atualizarInterface() {
    const metricas = calcularMetricas();
    console.log('📊 Métricas calculadas:', metricas);
    
    // Atualizar totais por área
    document.getElementById('totalMobileCliente').textContent = metricas.mobileCliente.global;
    document.getElementById('totalMobileDG').textContent = metricas.mobileDG.global;
    document.getElementById('totalInformatica').textContent = metricas.informatica.global;
    document.getElementById('totalPequenosDomesticos').textContent = metricas.pequenosDomesticos.global;
    document.getElementById('totalSomImagem').textContent = metricas.somImagem.global;
    document.getElementById('totalEntretenimento').textContent = metricas.entretenimento.global;
    
    // Para cada área, atualizar os novos campos
    // NOTA: Precisamos de adicionar novos elementos HTML para mostrar estes dados
    // Por enquanto, mostramos no console
    
    // Atualizar KPIs globais
    let totalGeral = 0;
    Object.values(metricas).forEach(area => totalGeral += area.global);
    document.getElementById('totalEntradas').textContent = totalGeral;
    
    console.log('✅ Interface atualizada com novas métricas');
}

// Dados de exemplo (fallback)
function carregarDadosExemplo() {
    dadosBrutos = [];
    const areas = ['Mobile Cliente', 'Mobile D&G', 'Informática', 'Pequenos Domésticos', 'Som e Imagem', 'Entretenimento'];
    const negocios = ['Garantias', 'Fora de Garantia', 'Extensão de Garantia'];
    
    for (let i = 1; i <= 200; i++) {
        const area = areas[Math.floor(Math.random() * areas.length)];
        let negocio;
        
        if (area === 'Mobile D&G') {
            negocio = 'D&G';
        } else {
            negocio = negocios[Math.floor(Math.random() * negocios.length)];
        }
        
        const data = new Date();
        data.setDate(data.getDate() - Math.floor(Math.random() * 30));
        
        dadosBrutos.push({
            id: i,
            area: area,
            negocio: negocio,
            status: Math.random() > 0.3 ? 'concluido' : (Math.random() > 0.5 ? 'andamento' : 'pendente'),
            data_entrada: data.toISOString().split('T')[0],
            tempo_reparo: Math.random() * 8,
            pendente_peca: Math.random() > 0.7,
            aguarda_cotacao: Math.random() > 0.8,
            is_orcamento: Math.random() > 0.9,
            polo: 'TSC South'
        });
    }
    
    ultimaAtualizacao = new Date();
    atualizarInterface();
    console.log('✅ Dados de exemplo carregados');
}
