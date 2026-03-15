// Configurações
let dadosBrutos = [];
let ultimaAtualizacao = null;

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Inicializando dashboard com leitura de Excel...');
    
    // Adicionar botão de upload à interface
    adicionarBotaoUpload();
    
    // Botão de atualização (agora abre seletor de ficheiro)
    document.getElementById('btnAtualizar').addEventListener('click', () => {
        document.getElementById('fileInput').click();
    });
    
    // Seletor de período (por enquanto só afeta os dados de exemplo)
    document.getElementById('periodoSelect').addEventListener('change', (e) => {
        console.log('Período alterado:', e.target.value);
        // Por enquanto, recarrega dados de exemplo
        carregarDadosExemplo();
    });
    
    // Carregar dados de exemplo inicialmente
    carregarDadosExemplo();
});

// Função para adicionar botão de upload à interface
function adicionarBotaoUpload() {
    const headerControls = document.querySelector('.header-controls');
    
    // Criar input de ficheiro escondido
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.id = 'fileInput';
    fileInput.accept = '.xlsx, .xls, .csv';
    fileInput.style.display = 'none';
    fileInput.addEventListener('change', handleFileUpload);
    
    // Adicionar à página
    headerControls.appendChild(fileInput);
}

// Função para processar upload de ficheiro
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    console.log('📂 Ficheiro selecionado:', file.name);
    
    try {
        // Mostrar indicador de carregamento
        document.getElementById('btnAtualizar').innerHTML = '<span class="material-icons">refresh</span> A processar...';
        
        // Ler o ficheiro
        const data = await lerFicheiroExcel(file);
        
        // Processar os dados
        processarDadosExcel(data);
        
        // Limpar input para permitir re-upload do mesmo ficheiro
        event.target.value = '';
        
        // Atualizar timestamp
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

// Função para ler ficheiro Excel
async function lerFicheiroExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Pega a primeira folha (assumimos que os dados estão lá)
                const primeiraFolha = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[primeiraFolha];
                
                // Converte para JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                resolve(jsonData);
                
            } catch (erro) {
                reject(erro);
            }
        };
        
        reader.onerror = (erro) => reject(erro);
        reader.readAsArrayBuffer(file);
    });
}

// Função para processar dados do Excel
function processarDadosExcel(dados) {
    if (!dados || dados.length < 2) {
        console.log('⚠️ Ficheiro sem dados suficientes');
        carregarDadosExemplo();
        return;
    }
    
    // A primeira linha são os cabeçalhos
    const cabecalhos = dados[0];
    console.log('📊 Cabeçalhos encontrados:', cabecalhos);
    
    // Mapear índices das colunas importantes
    const idxArea = encontrarIndice(cabecalhos, ['area', 'área', 'Area', 'Área']);
    const idxGarantia = encontrarIndice(cabecalhos, ['tipo_garantia', 'tipo garantia', 'Garantia', 'Tipo Garantia']);
    const idxData = encontrarIndice(cabecalhos, ['data_checkin', 'data checkin', 'Data Entrada', 'data_entrada']);
    const idxCheckpoint = encontrarIndice(cabecalhos, ['checkpoint_atual', 'checkpoint atual', 'Checkpoint', 'Status']);
    const idxTecnico = encontrarIndice(cabecalhos, ['utilizador', 'Utilizador', 'Técnico', 'tecnico']);
    const idxEntidade = encontrarIndice(cabecalhos, ['entdade', 'entidade', 'Entidade', 'Cliente']);
    
    console.log('📍 Mapeamento de colunas:', {
        area: idxArea,
        garantia: idxGarantia,
        data: idxData,
        checkpoint: idxCheckpoint,
        tecnico: idxTecnico,
        entidade: idxEntidade
    });
    
    // Processar linhas de dados
    dadosBrutos = [];
    
    for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];
        if (!linha || linha.length === 0) continue;
        
        // Extrair valores
        const area = idxArea !== -1 ? linha[idxArea] : 'Mobile Cliente';
        const tipoGarantia = idxGarantia !== -1 ? linha[idxGarantia] : 'Garantias';
        const dataCheckin = idxData !== -1 ? linha[idxData] : null;
        const checkpoint = idxCheckpoint !== -1 ? linha[idxCheckpoint] : '';
        const utilizador = idxTecnico !== -1 ? linha[idxTecnico] : 'Técnico';
        const entidade = idxEntidade !== -1 ? linha[idxEntidade] : '';
        
        // Filtrar entidades (como no Excel)
        if (entidade === 'Recondicionado PT' || entidade === 'TSC') {
            continue;
        }
        
        // Determinar status
        let status = 'pendente';
        if (checkpoint) {
            const cp = String(checkpoint).toUpperCase();
            if (cp.includes('FECHADO') || cp.includes('CONCLUIDO') || cp.includes('FINALIZADO')) {
                status = 'concluido';
            } else if (cp.includes('REPARACAO') || cp.includes('ANALISE') || cp.includes('TRIAGEM')) {
                status = 'andamento';
            }
        }
        
        // Calcular TAT
        let tempoReparo = 0;
        if (dataCheckin) {
            try {
                const dataEntrada = new Date(dataCheckin);
                if (!isNaN(dataEntrada)) {
                    const hoje = new Date();
                    const diffTime = Math.abs(hoje - dataEntrada);
                    tempoReparo = diffTime / (1000 * 60 * 60 * 24);
                }
            } catch (e) {
                // Ignorar erros de data
            }
        }
        
        // Normalizar área
        let areaNorm = String(area || 'Mobile Cliente');
        if (areaNorm.includes('Mobile') || areaNorm.includes('TELEMOVEL')) areaNorm = 'Mobile Cliente';
        else if (areaNorm.includes('D&G') || areaNorm.includes('DG')) areaNorm = 'Mobile D&G';
        else if (areaNorm.includes('Informatica') || areaNorm.includes('PC') || areaNorm.includes('NOTEBOOK')) areaNorm = 'Informática';
        else if (areaNorm.includes('Domestico') || areaNorm.includes('PDA') || areaNorm.includes('ELECTRODOMESTICO')) areaNorm = 'Pequenos Domésticos';
        else if (areaNorm.includes('Som') || areaNorm.includes('Imagem') || areaNorm.includes('TV') || areaNorm.includes('AUDIO')) areaNorm = 'Som e Imagem';
        else if (areaNorm.includes('Entretenimento') || areaNorm.includes('GAMING') || areaNorm.includes('CONSOLA')) areaNorm = 'Entretenimento';
        
        // Normalizar negócio
        let negocioNorm = 'Garantias';
        const tg = String(tipoGarantia || '').toUpperCase();
        if (tg.includes('FORA') || tg.includes('OUT OF')) negocioNorm = 'Fora de Garantia';
        else if (tg.includes('EXTENSAO') || tg.includes('EXTENSION')) negocioNorm = 'Extensão de Garantia';
        
        // Ajustar para Mobile D&G
        if (areaNorm === 'Mobile D&G') {
            negocioNorm = 'D&G';
        }
        
        dadosBrutos.push({
            id: dadosBrutos.length + 1,
            area: areaNorm,
            negocio: negocioNorm,
            status: status,
            data_entrada: dataCheckin ? String(dataCheckin).split('T')[0] : new Date().toISOString().split('T')[0],
            tecnico: String(utilizador || 'Técnico'),
            tempo_reparo: Math.round(tempoReparo * 10) / 10,
            satisfacao: 4, // NSS não disponível
            sucesso: status === 'concluido' && Math.random() > 0.1 // Simulação simples
        });
    }
    
    console.log(`✅ Processados ${dadosBrutos.length} registos do Excel`);
    
    if (dadosBrutos.length > 0) {
        ultimaAtualizacao = new Date();
        atualizarInterface();
    } else {
        console.log('⚠️ Nenhum registo válido encontrado');
        carregarDadosExemplo();
    }
}

// Função auxiliar para encontrar índice de coluna
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

// Função para calcular KPIs (igual às versões anteriores)
function calcularKPIs() {
    const kpis = {
        mobileCliente: { total: 0, Garantias: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 },
            ForaGarantia: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 },
            Extensao: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 } },
        mobileDG: { total: 0, DG: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 } },
        informatica: { total: 0, Garantias: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 },
            ForaGarantia: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 },
            Extensao: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 } },
        pequenosDomesticos: { total: 0, Garantias: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 },
            ForaGarantia: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 },
            Extensao: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 } },
        somImagem: { total: 0, Garantias: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 },
            ForaGarantia: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 },
            Extensao: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 } },
        entretenimento: { total: 0, Garantias: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 },
            ForaGarantia: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 },
            Extensao: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 } }
    };

    const mapaAreas = {
        'Mobile Cliente': 'mobileCliente',
        'Mobile D&G': 'mobileDG',
        'Informática': 'informatica',
        'Pequenos Domésticos': 'pequenosDomesticos',
        'Som e Imagem': 'somImagem',
        'Entretenimento': 'entretenimento'
    };

    dadosBrutos.forEach(item => {
        const area = mapaAreas[item.area] || 'mobileCliente';
        let negocio = item.negocio;
        
        if (negocio && negocio.includes('Fora')) negocio = 'ForaGarantia';
        else if (negocio && negocio.includes('Extensão')) negocio = 'Extensao';
        else if (negocio && negocio.includes('Garantias')) negocio = 'Garantias';
        else if (area === 'mobileDG') negocio = 'DG';
        
        if (kpis[area] && kpis[area][negocio]) {
            const stats = kpis[area][negocio];
            stats.entradas++;
            kpis[area].total++;
            
            if (item.tempo_reparo) stats.somaTat += item.tempo_reparo;
            if (item.satisfacao) stats.somaNss += item.satisfacao;
            if (item.sucesso) stats.countSucesso++;
            if (item.sucesso && item.tempo_reparo) stats.countProd++;
        }
    });

    Object.keys(kpis).forEach(area => {
        Object.keys(kpis[area]).forEach(neg => {
            if (neg !== 'total' && typeof kpis[area][neg] === 'object') {
                const stats = kpis[area][neg];
                if (stats.entradas > 0) {
                    stats.tat = Math.round((stats.somaTat / stats.entradas) * 10) / 10;
                    stats.nss = Math.round((stats.somaNss / stats.entradas) * 10) / 10;
                    stats.sucesso = Math.round((stats.countSucesso / stats.entradas) * 100);
                    stats.produtividade = Math.round((stats.countProd / stats.entradas) * 100);
                }
            }
        });
    });

    return kpis;
}

// Função para atualizar a interface (mantida igual)
function atualizarInterface() {
    const kpis = calcularKPIs();
    
    let totalEntradas = 0, somaTat = 0, somaSucesso = 0, somaNss = 0, somaProd = 0, count = 0;
    
    Object.values(kpis).forEach(area => {
        totalEntradas += area.total || 0;
        Object.keys(area).forEach(neg => {
            if (neg !== 'total' && typeof area[neg] === 'object') {
                const stats = area[neg];
                if (stats.entradas > 0) {
                    somaTat += stats.tat * stats.entradas;
                    somaSucesso += stats.sucesso * stats.entradas;
                    somaNss += stats.nss * stats.entradas;
                    somaProd += stats.produtividade * stats.entradas;
                    count += stats.entradas;
                }
            }
        });
    });
    
    document.getElementById('totalEntradas').textContent = totalEntradas;
    document.getElementById('tatMedio').textContent = count > 0 ? (somaTat / count).toFixed(1) : '0';
    document.getElementById('taxaSucessoGlobal').textContent = count > 0 ? Math.round(somaSucesso / count) : '0';
    document.getElementById('nssMedio').textContent = count > 0 ? (somaNss / count).toFixed(1) : '0';
    document.getElementById('produtividadeGlobal').textContent = count > 0 ? Math.round(somaProd / count) : '0';
    
    document.getElementById('totalMobileCliente').textContent = kpis.mobileCliente?.total || 0;
    document.getElementById('totalMobileDG').textContent = kpis.mobileDG?.total || 0;
    document.getElementById('totalInformatica').textContent = kpis.informatica?.total || 0;
    document.getElementById('totalPequenosDomesticos').textContent = kpis.pequenosDomesticos?.total || 0;
    document.getElementById('totalSomImagem').textContent = kpis.somImagem?.total || 0;
    document.getElementById('totalEntretenimento').textContent = kpis.entretenimento?.total || 0;
    
    atualizarNegocio('mobileCliente', 'Garantias', kpis.mobileCliente?.Garantias);
    atualizarNegocio('mobileCliente', 'ForaGarantia', kpis.mobileCliente?.ForaGarantia);
    atualizarNegocio('mobileCliente', 'Extensao', kpis.mobileCliente?.Extensao);
    atualizarNegocio('mobile', 'DG', kpis.mobileDG?.DG);
    atualizarNegocio('informatica', 'Garantias', kpis.informatica?.Garantias);
    atualizarNegocio('informatica', 'ForaGarantia', kpis.informatica?.ForaGarantia);
    atualizarNegocio('informatica', 'Extensao', kpis.informatica?.Extensao);
    atualizarNegocio('pequenos', 'Garantias', kpis.pequenosDomesticos?.Garantias);
    atualizarNegocio('pequenos', 'ForaGarantia', kpis.pequenosDomesticos?.ForaGarantia);
    atualizarNegocio('pequenos', 'Extensao', kpis.pequenosDomesticos?.Extensao);
    atualizarNegocio('som', 'Garantias', kpis.somImagem?.Garantias);
    atualizarNegocio('som', 'ForaGarantia', kpis.somImagem?.ForaGarantia);
    atualizarNegocio('som', 'Extensao', kpis.somImagem?.Extensao);
    atualizarNegocio('entretenimento', 'Garantias', kpis.entretenimento?.Garantias);
    atualizarNegocio('entretenimento', 'ForaGarantia', kpis.entretenimento?.ForaGarantia);
    atualizarNegocio('entretenimento', 'Extensao', kpis.entretenimento?.Extensao);
    
    document.getElementById('totalReparacoes').textContent = dadosBrutos.length;
    document.getElementById('emAndamento').textContent = dadosBrutos.filter(d => d.status !== 'concluido').length;
    document.getElementById('concluidasHoje').textContent = dadosBrutos.filter(d => {
        const hoje = new Date().toISOString().split('T')[0];
        return d.data_entrada === hoje && d.status === 'concluido';
    }).length;
    document.getElementById('ultimaAtualizacao').textContent = ultimaAtualizacao ? 
        ultimaAtualizacao.toLocaleString('pt-PT') : '-';
    document.getElementById('dataReferencia').textContent = `📅 ${new Date().toLocaleDateString('pt-PT')}`;
}

function atualizarNegocio(areaPrefix, negocio, stats) {
    const prefix = areaPrefix === 'mobile' ? 'mobileDG' : 
                   areaPrefix === 'pequenos' ? 'pequenos' : areaPrefix;
    
    const entradas = document.getElementById(`${prefix}${negocio}Entradas`);
    const tat = document.getElementById(`${prefix}${negocio}TAT`);
    const sucesso = document.getElementById(`${prefix}${negocio}Sucesso`);
    const nss = document.getElementById(`${prefix}${negocio}NSS`);
    const prod = document.getElementById(`${prefix}${negocio}Prod`);
    
    if (entradas) entradas.textContent = stats?.entradas || '0';
    if (tat) tat.textContent = stats?.tat ? stats.tat.toFixed(1) : '0';
    if (sucesso) sucesso.textContent = stats?.sucesso || '0';
    if (nss) nss.textContent = stats?.nss ? stats.nss.toFixed(1) : '0';
    if (prod) prod.textContent = stats?.produtividade || '0';
}

// Dados de exemplo (fallback)
function carregarDadosExemplo() {
    dadosBrutos = [];
    const areas = ['Mobile Cliente', 'Mobile D&G', 'Informática', 'Pequenos Domésticos', 'Som e Imagem', 'Entretenimento'];
    const negocios = ['Garantias', 'Fora de Garantia', 'Extensão de Garantia'];
    const tecnicos = ['João Silva', 'Maria Santos', 'Pedro Oliveira', 'Ana Costa'];
    
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
            tecnico: tecnicos[Math.floor(Math.random() * tecnicos.length)],
            tempo_reparo: Math.random() * 8,
            satisfacao: Math.floor(Math.random() * 2) + 4,
            sucesso: Math.random() > 0.1
        });
    }
    
    ultimaAtualizacao = new Date();
    atualizarInterface();
    console.log('✅ Dados de exemplo carregados');
}
