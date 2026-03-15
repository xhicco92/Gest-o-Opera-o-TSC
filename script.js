// Configurações
let dadosBrutos = [];
let ultimaAtualizacao = null;
let cabecalhosOriginais = [];

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Inicializando dashboard com novos indicadores...');
    
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
    
    // Mapear índices
    const idxPolo = encontrarIndice(cabecalhos, ['polo', 'Polo', 'unidade']);
    const idxArea = encontrarIndice(cabecalhos, ['area', 'Área', 'departamento']);
    const idxGarantia = encontrarIndice(cabecalhos, ['tipo_garantia', 'garantia', 'Tipo Garantia']);
    const idxData = encontrarIndice(cabecalhos, ['data_checkin', 'checkin', 'Data Entrada']);
    const idxPendentePeca = encontrarIndice(cabecalhos, ['pendente_peca', 'Pendente Peça', 'aguarda peça']);
    const idxAguardaCotacao = encontrarIndice(cabecalhos, ['aguarda_cotacao_de_peca', 'aguarda cotação', 'cotação']);
    const idxTipoReparacao = encontrarIndice(cabecalhos, ['tipo_reparacao', 'Tipo Reparação']);
    
    console.log('📍 Mapeamento:', {
        polo: idxPolo,
        area: idxArea,
        garantia: idxGarantia,
        data: idxData,
        pendente_peca: idxPendentePeca,
        aguarda_cotacao: idxAguardaCotacao,
        tipo_reparacao: idxTipoReparacao
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
        const area = idxArea !== -1 ? linha[idxArea] : '';
        const tipoGarantia = idxGarantia !== -1 ? linha[idxGarantia] : '';
        const dataCheckin = idxData !== -1 ? linha[idxData] : null;
        const pendentePeca = idxPendentePeca !== -1 ? String(linha[idxPendentePeca] || '').toLowerCase() : '';
        const aguardaCotacao = idxAguardaCotacao !== -1 ? String(linha[idxAguardaCotacao] || '').toLowerCase() : '';
        const tipoReparacao = idxTipoReparacao !== -1 ? String(linha[idxTipoReparacao] || '').toLowerCase() : '';
        
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
        
        // Normalizar área
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
            data_entrada: dataCheckin,
            tempo_reparo: Math.round(tempoReparo * 10) / 10,
            pendente_peca: isPendentePeca,
            aguarda_cotacao: isAguardaCotacao,
            is_orcamento: isOrcamento,
            tipo_garantia: tipoGarantia,
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

// Função para calcular as métricas por área e negócio
function calcularMetricas() {
    // Estrutura para cada combinação área + negócio
    const metricas = {
        mobileCliente: {
            Garantias: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 },
            ForaGarantia: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 },
            Extensao: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 }
        },
        mobileDG: {
            DG: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 }
        },
        informatica: {
            Garantias: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 },
            ForaGarantia: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 },
            Extensao: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 }
        },
        pequenosDomesticos: {
            Garantias: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 },
            ForaGarantia: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 },
            Extensao: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 }
        },
        somImagem: {
            Garantias: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 },
            ForaGarantia: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 },
            Extensao: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 }
        },
        entretenimento: {
            Garantias: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 },
            ForaGarantia: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 },
            Extensao: { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0, somaTat:0, countTat:0 }
        }
    };

    const mapaAreas = {
        'Mobile Cliente': 'mobileCliente',
        'Mobile D&G': 'mobileDG',
        'Informática': 'informatica',
        'Pequenos Domésticos': 'pequenosDomesticos',
        'Som e Imagem': 'somImagem',
        'Entretenimento': 'entretenimento'
    };

    // Processar cada registo
    dadosBrutos.forEach(item => {
        const areaKey = mapaAreas[item.area];
        if (!areaKey) return;
        
        let negocioKey;
        if (item.negocio === 'Garantias') negocioKey = 'Garantias';
        else if (item.negocio === 'Fora de Garantia') negocioKey = 'ForaGarantia';
        else if (item.negocio === 'Extensão de Garantia') negocioKey = 'Extensao';
        else if (item.negocio === 'D&G') negocioKey = 'DG';
        else return;
        
        const stats = metricas[areaKey][negocioKey];
        if (!stats) return;
        
        // Análises (contagem por tipo de garantia)
        stats.analises++;
        
        // Reparações (apenas não concluídos)
        if (item.tempo_reparo > 0) { // Consideramos que tempo_reparo > 0 significa em aberto
            stats.reparacoesGlobal++;
            if (item.pendente_peca) {
                stats.reparacoesPendentes++;
            } else {
                stats.reparacoesNaoPendentes++;
            }
        }
        
        // TAT Aberto
        if (item.tempo_reparo > 0) {
            stats.somaTat += item.tempo_reparo;
            stats.countTat++;
        }
        
        // Orçamentos
        if (item.is_orcamento) {
            stats.orcamentos++;
        }
        
        // Aguarda Aceitação
        if (item.aguarda_cotacao) {
            stats.aguardaAceitacao++;
        }
        
        // Débitos WSS (todos os registos por enquanto)
        stats.debitosWSS = stats.analises; // Depois ajustamos conforme definição
    });

    // Calcular médias
    Object.keys(metricas).forEach(area => {
        Object.keys(metricas[area]).forEach(neg => {
            const stats = metricas[area][neg];
            if (stats.countTat > 0) {
                stats.tatAberto = Math.round((stats.somaTat / stats.countTat) * 10) / 10;
            }
        });
    });

    return metricas;
}

// Função para atualizar a interface com os novos campos
function atualizarInterface() {
    const metricas = calcularMetricas();
    console.log('📊 Métricas calculadas:', metricas);
    
    // Totais globais (somando todas as áreas)
    let totalGeral = 0;
    Object.values(metricas).forEach(area => {
        Object.values(area).forEach(neg => {
            if (typeof neg === 'object') {
                totalGeral += neg.analises || 0;
            }
        });
    });
    document.getElementById('totalEntradas').textContent = totalGeral;
    
    // TAT Médio Global
    let somaTatGlobal = 0, countTatGlobal = 0;
    Object.values(metricas).forEach(area => {
        Object.values(area).forEach(neg => {
            if (typeof neg === 'object') {
                somaTatGlobal += neg.somaTat || 0;
                countTatGlobal += neg.countTat || 0;
            }
        });
    });
    const tatMedio = countTatGlobal > 0 ? (somaTatGlobal / countTatGlobal).toFixed(1) : 0;
    document.getElementById('tatMedio').textContent = tatMedio;
    
    // Taxa de Sucesso Global (reparações não pendentes / total reparações)
    let reparacoesTotal = 0, reparacoesNaoPendentes = 0;
    Object.values(metricas).forEach(area => {
        Object.values(area).forEach(neg => {
            if (typeof neg === 'object') {
                reparacoesTotal += neg.reparacoesGlobal || 0;
                reparacoesNaoPendentes += neg.reparacoesNaoPendentes || 0;
            }
        });
    });
    const taxaSucesso = reparacoesTotal > 0 ? Math.round((reparacoesNaoPendentes / reparacoesTotal) * 100) : 0;
    document.getElementById('taxaSucessoGlobal').textContent = taxaSucesso;
    
    // NSS Global (substituído por Orçamentos + Aguarda Aceitação)
    let orcamentosTotal = 0, aguardaTotal = 0;
    Object.values(metricas).forEach(area => {
        Object.values(area).forEach(neg => {
            if (typeof neg === 'object') {
                orcamentosTotal += neg.orcamentos || 0;
                aguardaTotal += neg.aguardaAceitacao || 0;
            }
        });
    });
    document.getElementById('nssMedio').textContent = orcamentosTotal; // Temporário
    document.getElementById('produtividadeGlobal').textContent = aguardaTotal; // Temporário
    
    // Atualizar totais por área (mantém os IDs existentes)
    document.getElementById('totalMobileCliente').textContent = 
        (metricas.mobileCliente?.Garantias?.analises || 0) + 
        (metricas.mobileCliente?.ForaGarantia?.analises || 0) + 
        (metricas.mobileCliente?.Extensao?.analises || 0);
    
    document.getElementById('totalMobileDG').textContent = metricas.mobileDG?.DG?.analises || 0;
    
    document.getElementById('totalInformatica').textContent = 
        (metricas.informatica?.Garantias?.analises || 0) + 
        (metricas.informatica?.ForaGarantia?.analises || 0) + 
        (metricas.informatica?.Extensao?.analises || 0);
    
    document.getElementById('totalPequenosDomesticos').textContent = 
        (metricas.pequenosDomesticos?.Garantias?.analises || 0) + 
        (metricas.pequenosDomesticos?.ForaGarantia?.analises || 0) + 
        (metricas.pequenosDomesticos?.Extensao?.analises || 0);
    
    document.getElementById('totalSomImagem').textContent = 
        (metricas.somImagem?.Garantias?.analises || 0) + 
        (metricas.somImagem?.ForaGarantia?.analises || 0) + 
        (metricas.somImagem?.Extensao?.analises || 0);
    
    document.getElementById('totalEntretenimento').textContent = 
        (metricas.entretenimento?.Garantias?.analises || 0) + 
        (metricas.entretenimento?.ForaGarantia?.analises || 0) + 
        (metricas.entretenimento?.Extensao?.analises || 0);
    
    // Agora atualizar os campos de cada negócio com os novos indicadores
    // Mobile Cliente
    atualizarNegocioMetricas('mobileCliente', 'Garantias', metricas.mobileCliente?.Garantias);
    atualizarNegocioMetricas('mobileCliente', 'ForaGarantia', metricas.mobileCliente?.ForaGarantia);
    atualizarNegocioMetricas('mobileCliente', 'Extensao', metricas.mobileCliente?.Extensao);
    
    // Mobile D&G
    atualizarNegocioMetricas('mobile', 'DG', metricas.mobileDG?.DG);
    
    // Informática
    atualizarNegocioMetricas('informatica', 'Garantias', metricas.informatica?.Garantias);
    atualizarNegocioMetricas('informatica', 'ForaGarantia', metricas.informatica?.ForaGarantia);
    atualizarNegocioMetricas('informatica', 'Extensao', metricas.informatica?.Extensao);
    
    // Pequenos Domésticos
    atualizarNegocioMetricas('pequenos', 'Garantias', metricas.pequenosDomesticos?.Garantias);
    atualizarNegocioMetricas('pequenos', 'ForaGarantia', metricas.pequenosDomesticos?.ForaGarantia);
    atualizarNegocioMetricas('pequenos', 'Extensao', metricas.pequenosDomesticos?.Extensao);
    
    // Som e Imagem
    atualizarNegocioMetricas('som', 'Garantias', metricas.somImagem?.Garantias);
    atualizarNegocioMetricas('som', 'ForaGarantia', metricas.somImagem?.ForaGarantia);
    atualizarNegocioMetricas('som', 'Extensao', metricas.somImagem?.Extensao);
    
    // Entretenimento
    atualizarNegocioMetricas('entretenimento', 'Garantias', metricas.entretenimento?.Garantias);
    atualizarNegocioMetricas('entretenimento', 'ForaGarantia', metricas.entretenimento?.ForaGarantia);
    atualizarNegocioMetricas('entretenimento', 'Extensao', metricas.entretenimento?.Extensao);
    
    // Rodapé
    document.getElementById('totalReparacoes').textContent = dadosBrutos.length;
    document.getElementById('emAndamento').textContent = dadosBrutos.filter(d => d.tempo_reparo > 0).length;
    document.getElementById('concluidasHoje').textContent = dadosBrutos.filter(d => {
        if (!d.data_entrada) return false;
        const hoje = new Date().toISOString().split('T')[0];
        const dataStr = String(d.data_entrada).split('T')[0];
        return dataStr === hoje;
    }).length;
    document.getElementById('ultimaAtualizacao').textContent = ultimaAtualizacao ? 
        ultimaAtualizacao.toLocaleString('pt-PT') : '-';
    document.getElementById('dataReferencia').textContent = `📅 ${new Date().toLocaleDateString('pt-PT')}`;
}

// Função para atualizar um negócio com as novas métricas
function atualizarNegocioMetricas(areaPrefix, negocio, stats) {
    const prefix = areaPrefix === 'mobile' ? 'mobileDG' : 
                   areaPrefix === 'pequenos' ? 'pequenos' : areaPrefix;
    
    if (!stats) stats = { analises:0, reparacoesGlobal:0, reparacoesPendentes:0, reparacoesNaoPendentes:0, tatAberto:0, orcamentos:0, aguardaAceitacao:0, debitosWSS:0 };
    
    // Mapear para os IDs existentes no HTML
    // NOTA: Os IDs no HTML são: entradas, TAT, sucesso, NSS, prod
    // Vamos reutilizá-los com os novos significados
    
    const entradas = document.getElementById(`${prefix}${negocio}Entradas`);
    const tat = document.getElementById(`${prefix}${negocio}TAT`);
    const sucesso = document.getElementById(`${prefix}${negocio}Sucesso`);
    const nss = document.getElementById(`${prefix}${negocio}NSS`);
    const prod = document.getElementById(`${prefix}${negocio}Prod`);
    
    if (entradas) {
        // Análises
        entradas.textContent = stats.analises || 0;
        entradas.parentElement.querySelector('.mini-label').textContent = 'Análises';
    }
    
    if (tat) {
        // Reparações Global
        tat.textContent = stats.reparacoesGlobal || 0;
        tat.parentElement.querySelector('.mini-label').textContent = 'Rep.Global';
    }
    
    if (sucesso) {
        // Pendentes Peça
        sucesso.textContent = stats.reparacoesPendentes || 0;
        sucesso.parentElement.querySelector('.mini-label').textContent = 'Pend.Peça';
    }
    
    if (nss) {
        // TAT Aberto
        nss.textContent = stats.tatAberto ? stats.tatAberto.toFixed(1) : '0';
        nss.parentElement.querySelector('.mini-label').textContent = 'TAT Aberto';
    }
    
    if (prod) {
        // Orçamentos + Aguarda (temporário)
        prod.textContent = (stats.orcamentos || 0) + ' | ' + (stats.aguardaAceitacao || 0);
        prod.parentElement.querySelector('.mini-label').textContent = 'Orç|Aguarda';
    }
}

// Dados de exemplo
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

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Inicializando dashboard...');
    
    document.getElementById('btnAtualizar').addEventListener('click', () => {
        document.getElementById('btnAtualizar').innerHTML = '<span class="material-icons">refresh</span> A atualizar...';
        document.getElementById('fileInput').click();
    });
    
    document.getElementById('periodoSelect').addEventListener('change', (e) => {
        console.log('Período alterado:', e.target.value);
        carregarDadosExemplo();
    });
    
    carregarDadosExemplo();
});
