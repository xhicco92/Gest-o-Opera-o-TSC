// Configurações
let dadosBrutos = [];
let ultimaAtualizacao = null;
let cabecalhosOriginais = [];

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Inicializando dashboard com leitura da folha "Dados"...');
    
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
                
                // Verificar se a folha "Dados" existe
                console.log('📑 Folhas disponíveis:', workbook.SheetNames);
                
                if (!workbook.SheetNames.includes('Dados')) {
                    console.log('⚠️ Folha "Dados" não encontrada. Usando primeira folha:', workbook.SheetNames[0]);
                }
                
                // Usar folha "Dados" ou a primeira
                const nomeFolha = workbook.SheetNames.includes('Dados') ? 'Dados' : workbook.SheetNames[0];
                const worksheet = workbook.Sheets[nomeFolha];
                
                // Converter para JSON mantendo cabeçalhos
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1,
                    defval: '' // Valor padrão para células vazias
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
    
    // Primeira linha são os cabeçalhos
    const cabecalhos = dados[0];
    cabecalhosOriginais = cabecalhos;
    console.log('📊 Cabeçalhos encontrados:', cabecalhos);
    
    // Mapear índices das colunas importantes (baseado nos nomes reais do Excel)
    const idxArea = encontrarIndice(cabecalhos, ['area', 'área', 'Area', 'Área', 'polo']);
    const idxGarantia = encontrarIndice(cabecalhos, ['tipo_garantia', 'tipo garantia', 'Garantia', 'Tipo Garantia', 'garantia']);
    const idxData = encontrarIndice(cabecalhos, ['data_checkin', 'data checkin', 'Data Entrada', 'data_entrada', 'checkin']);
    const idxCheckpoint = encontrarIndice(cabecalhos, ['checkpoint_atual', 'checkpoint atual', 'Checkpoint', 'Status', 'estado']);
    const idxTecnico = encontrarIndice(cabecalhos, ['utilizador', 'Utilizador', 'Técnico', 'tecnico', 'responsavel']);
    const idxEntidade = encontrarIndice(cabecalhos, ['entdade', 'entidade', 'Entidade', 'Cliente', 'entidade']);
    const idxMarca = encontrarIndice(cabecalhos, ['marca', 'Marca', 'brand', 'marca']);
    const idxModelo = encontrarIndice(cabecalhos, ['modelo', 'Modelo', 'model']);
    const idxSn = encontrarIndice(cabecalhos, ['sn_equipamento', 'sn', 'número série', 'serial', 'imei']);
    const idxId = encontrarIndice(cabecalhos, ['id_processo', 'id', 'processo', 'os', 'ordem']);
    
    console.log('📍 Mapeamento de colunas:', {
        area: idxArea,
        garantia: idxGarantia,
        data: idxData,
        checkpoint: idxCheckpoint,
        tecnico: idxTecnico,
        entidade: idxEntidade,
        marca: idxMarca,
        modelo: idxModelo,
        sn: idxSn,
        id: idxId
    });
    
    // Processar linhas de dados (a partir da linha 2)
    dadosBrutos = [];
    
    for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];
        if (!linha || linha.length === 0) continue;
        
        // Verificar se a linha tem conteúdo relevante
        if (linha.every(cell => !cell || cell === '')) continue;
        
        // Extrair valores
        const area = idxArea !== -1 ? linha[idxArea] : '';
        const tipoGarantia = idxGarantia !== -1 ? linha[idxGarantia] : '';
        const dataCheckin = idxData !== -1 ? linha[idxData] : null;
        const checkpoint = idxCheckpoint !== -1 ? linha[idxCheckpoint] : '';
        const utilizador = idxTecnico !== -1 ? linha[idxTecnico] : '';
        const entidade = idxEntidade !== -1 ? linha[idxEntidade] : '';
        const marca = idxMarca !== -1 ? linha[idxMarca] : '';
        const modelo = idxModelo !== -1 ? linha[idxModelo] : '';
        const sn = idxSn !== -1 ? linha[idxSn] : '';
        const id = idxId !== -1 ? linha[idxId] : i;
        
        // Determinar status baseado no checkpoint
        let status = 'pendente';
        if (checkpoint) {
            const cp = String(checkpoint).toUpperCase();
            if (cp.includes('FECHADO') || cp.includes('CONCLUIDO') || cp.includes('FINALIZADO')) {
                status = 'concluido';
            } else if (cp.includes('REPARACAO') || cp.includes('ANALISE') || cp.includes('TRIAGEM')) {
                status = 'andamento';
            }
        }
        
        // Calcular TAT (dias desde data_checkin até hoje)
        let tempoReparo = 0;
        if (dataCheckin) {
            try {
                // Converter data do Excel (número ou string)
                let dataEntrada;
                if (typeof dataCheckin === 'number') {
                    // Data do Excel (dias desde 1900)
                    dataEntrada = new Date((dataCheckin - 25569) * 86400 * 1000);
                } else {
                    dataEntrada = new Date(dataCheckin);
                }
                
                if (!isNaN(dataEntrada)) {
                    const hoje = new Date();
                    const diffTime = Math.abs(hoje - dataEntrada);
                    tempoReparo = diffTime / (1000 * 60 * 60 * 24);
                }
            } catch (e) {
                console.log('Erro ao processar data:', dataCheckin);
            }
        }
        
        // Normalizar área
        let areaNorm = 'Outros';
        const areaStr = String(area || '').toUpperCase();
        
        if (areaStr.includes('MOBILE') || areaStr.includes('TELEMOVEL') || areaStr.includes('CELULAR')) {
            if (areaStr.includes('D&G') || areaStr.includes('DG')) {
                areaNorm = 'Mobile D&G';
            } else {
                areaNorm = 'Mobile Cliente';
            }
        } else if (areaStr.includes('INFORMATICA') || areaStr.includes('PC') || areaStr.includes('NOTEBOOK') || areaStr.includes('COMPUTADOR')) {
            areaNorm = 'Informática';
        } else if (areaStr.includes('DOMESTICO') || areaStr.includes('PDA') || areaStr.includes('ELECTRODOMESTICO') || areaStr.includes('PEQUENO')) {
            areaNorm = 'Pequenos Domésticos';
        } else if (areaStr.includes('SOM') || areaStr.includes('IMAGEM') || areaStr.includes('TV') || areaStr.includes('AUDIO') || areaStr.includes('VIDEO')) {
            areaNorm = 'Som e Imagem';
        } else if (areaStr.includes('ENTRETENIMENTO') || areaStr.includes('GAMING') || areaStr.includes('CONSOLA') || areaStr.includes('PLAYSTATION')) {
            areaNorm = 'Entretenimento';
        } else if (areaStr) {
            // Se tem valor mas não reconheceu, mantém original
            areaNorm = String(area);
        }
        
        // Normalizar negócio
        let negocioNorm = 'Garantias';
        const tg = String(tipoGarantia || '').toUpperCase();
        
        if (tg.includes('FORA') || tg.includes('OUT OF') || tg.includes('PAGO') || tg.includes('NAO GARANTIA')) {
            negocioNorm = 'Fora de Garantia';
        } else if (tg.includes('EXTENSAO') || tg.includes('EXTENSION') || tg.includes('PROTECAO')) {
            negocioNorm = 'Extensão de Garantia';
        }
        
        // Ajustar para Mobile D&G
        if (areaNorm === 'Mobile D&G') {
            negocioNorm = 'D&G';
        }
        
        // Determinar sucesso (aleatório por enquanto - depois podemos ajustar)
        const sucesso = status === 'concluido' ? Math.random() > 0.1 : true;
        
        // Criar objeto com todos os dados
        dadosBrutos.push({
            id: dadosBrutos.length + 1,
            id_original: id,
            area: areaNorm,
            area_original: area,
            negocio: negocioNorm,
            negocio_original: tipoGarantia,
            status: status,
            data_entrada: dataCheckin ? String(dataCheckin) : '',
            tecnico: String(utilizador || 'Não atribuído'),
            tempo_reparo: Math.round(tempoReparo * 10) / 10,
            satisfacao: 4, // NSS não disponível
            sucesso: sucesso,
            checkpoint: checkpoint,
            marca: String(marca || ''),
            modelo: String(modelo || ''),
            sn: String(sn || ''),
            entidade: String(entidade || '')
        });
    }
    
    console.log(`✅ Processados ${dadosBrutos.length} registos da folha "Dados"`);
    console.log('📊 Primeiro registo processado:', dadosBrutos[0]);
    
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
                console.log(`✅ Coluna "${cabecalhos[i]}" corresponde a "${nome}" (índice ${i})`);
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
            Extensao: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 } },
        outros: { total: 0, Outros: { entradas:0,tat:0,sucesso:0,nss:0,produtividade:0,somaTat:0,somaNss:0,countSucesso:0,countProd:0 } }
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

    dadosBrutos.forEach(item => {
        const area = mapaAreas[item.area] || 'outros';
        let negocio = item.negocio;
        
        if (negocio && negocio.includes('Fora')) negocio = 'ForaGarantia';
        else if (negocio && negocio.includes('Extensão')) negocio = 'Extensao';
        else if (negocio && negocio.includes('Garantias')) negocio = 'Garantias';
        else if (area === 'mobileDG') negocio = 'DG';
        else if (area === 'outros') negocio = 'Outros';
        
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

// Função para atualizar a interface
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

// Dados de exemplo
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
