// CONFIGURAÇÕES
const USERNAME = 'francisco.moreira@worten.pt';
const PASSWORD = 'Alice311020***';

// Estado da aplicação
let dadosBrutos = [];
let ultimaAtualizacao = null;

// Função para buscar dados da API
async function buscarDadosAPI() {
    try {
        console.log('🔄 Buscando dados da API com campos reais...');
        
        // Prepara os parâmetros (formato ISO que o Excel usa)
        const dataFim = new Date();
        const dataIni = new Date();
        dataIni.setDate(dataIni.getDate() - 30);
        
        const dataIniStr = dataIni.toISOString().split('T')[0];
        const dataFimStr = dataFim.toISOString().split('T')[0];
        
        const corpo = new URLSearchParams();
        corpo.append('UserName', USERNAME);
        corpo.append('Password', PASSWORD);
        corpo.append('dataIni', dataIniStr);
        corpo.append('dataFim', dataFimStr);
        
        // Headers iguais aos do Excel
        const headers = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept': 'application/xml, text/xml, */*',
            'Accept-Language': 'pt-PT,pt;q=0.9,en;q=0.8',
            'Cache-Control': 'no-cache',
            'User-Agent': 'Microsoft Excel/16.0'
        };
        
        console.log('📤 A enviar requisição...');
        
        // Usa o proxy cors-anywhere (precisa de ativação única)
        const proxyURL = 'https://cors-anywhere.herokuapp.com/';
        const resposta = await fetch(proxyURL + 'https://reportingwss.noshape.com/ServiceV3.asmx/TSC_OrsAbertasEmCadaCheckpoint_2', {
            method: 'POST',
            headers: headers,
            body: corpo.toString()
        });
        
        if (!resposta.ok) {
            throw new Error(`HTTP ${resposta.status}`);
        }
        
        const xmlText = await resposta.text();
        console.log('✅ Resposta recebida!');
        
        // Processa o XML
        processarXML(xmlText);
        
    } catch (erro) {
        console.error('❌ Erro ao buscar dados:', erro);
        
        // Mostra instruções para ativar o proxy
        if (!document.querySelector('.proxy-aviso')) {
            const aviso = document.createElement('div');
            aviso.className = 'proxy-aviso';
            aviso.innerHTML = '⚠️ Para aceder à API, precisa ativar o proxy temporário. ' +
                '<a href="https://cors-anywhere.herokuapp.com/corsdemo" target="_blank" style="color:#007bff;">Clique aqui</a>, ' +
                'peça acesso temporário e depois atualize a página.';
            aviso.style.cssText = 'background: #cce5ff; color: #004085; padding: 15px; border-radius: 5px; margin: 10px 0; border: 1px solid #b8daff; text-align:center;';
            document.querySelector('.dashboard-header').appendChild(aviso);
        }
        
        carregarDadosExemplo();
    }
}

// Função para processar o XML com os campos reais
function processarXML(xmlText) {
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, 'text/xml');
        
        // Verifica erros de parse
        const parseError = xmlDoc.querySelector('parsererror');
        if (parseError) {
            throw new Error('Erro ao parsear XML');
        }
        
        // Procura pelos registos (provavelmente em <Table> ou <Record>)
        let registos = [];
        const possiveisTags = ['Table', 'Record', 'row', 'Result'];
        
        for (let tag of possiveisTags) {
            const elementos = xmlDoc.getElementsByTagName(tag);
            if (elementos.length > 0) {
                console.log(`✅ Encontrados ${elementos.length} registos em <${tag}>`);
                registos = Array.from(elementos);
                break;
            }
        }
        
        if (registos.length === 0) {
            console.log('⚠️ Nenhum registo encontrado no XML');
            carregarDadosExemplo();
            return;
        }
        
        // Converte os registos para o formato do dashboard
        dadosBrutos = registos.map((reg, index) => {
            // Função auxiliar para extrair valor de campo
            const getValor = (nomeCampo) => {
                // Tenta diferentes variações do nome do campo
                const variacoes = [
                    reg.querySelector(nomeCampo),
                    reg.querySelector(nomeCampo.toLowerCase()),
                    reg.querySelector(nomeCampo.toUpperCase()),
                    reg.querySelector(nomeCampo.charAt(0).toUpperCase() + nomeCampo.slice(1))
                ];
                
                for (let elem of variacoes) {
                    if (elem) return elem.textContent;
                }
                return null;
            };
            
            // Mapeamento baseado nos campos do Excel
            const area = getValor('area') || getValor('Area') || 'Mobile Cliente';
            const tipoGarantia = getValor('tipo_garantia') || getValor('tipoGarantia') || 'Garantias';
            const dataCheckin = getValor('data_checkin') || getValor('dataCheckin');
            const dataCheckpoint = getValor('data_checkpoint_atual') || getValor('dataCheckpointAtual');
            const checkpoint = getValor('checkpoint_atual') || getValor('checkpointAtual');
            const utilizador = getValor('utilizador') || getValor('Utilizador') || 'Técnico';
            const pendentePeca = getValor('pendente_peca') || getValor('pendentePeca');
            
            // Determina o status com base no checkpoint
            let status = 'pendente';
            if (checkpoint) {
                if (checkpoint.includes('FECHADO') || checkpoint.includes('CONCLUIDO')) status = 'concluido';
                else if (checkpoint.includes('REPARACAO') || checkpoint.includes('ANALISE')) status = 'andamento';
            }
            
            // Calcula tempo de reparo (TAT) em dias
            let tempoReparo = 0;
            if (dataCheckin) {
                const dataEntrada = new Date(dataCheckin);
                const hoje = new Date();
                const diffTime = Math.abs(hoje - dataEntrada);
                tempoReparo = diffTime / (1000 * 60 * 60 * 24);
            }
            
            // Normaliza o nome da área
            let areaNormalizada = area;
            if (area.includes('Mobile') || area.includes('TELEMOVEL')) areaNormalizada = 'Mobile Cliente';
            else if (area.includes('D&G') || area.includes('DG')) areaNormalizada = 'Mobile D&G';
            else if (area.includes('Informatica') || area.includes('PC')) areaNormalizada = 'Informática';
            else if (area.includes('Domestico') || area.includes('PDA')) areaNormalizada = 'Pequenos Domésticos';
            else if (area.includes('Som') || area.includes('Imagem') || area.includes('TV')) areaNormalizada = 'Som e Imagem';
            else if (area.includes('Entretenimento') || area.includes('GAMING')) areaNormalizada = 'Entretenimento';
            
            // Normaliza o tipo de garantia
            let negocioNormalizado = tipoGarantia;
            if (tipoGarantia.includes('Garantia') && !tipoGarantia.includes('Fora')) negocioNormalizado = 'Garantias';
            else if (tipoGarantia.includes('Fora')) negocioNormalizado = 'Fora de Garantia';
            else if (tipoGarantia.includes('Extensao') || tipoGarantia.includes('Extensão')) negocioNormalizado = 'Extensão de Garantia';
            
            return {
                id: index + 1,
                area: areaNormalizada,
                negocio: negocioNormalizado,
                status: status,
                data_entrada: dataCheckin ? new Date(dataCheckin).toISOString().split('T')[0] : new Date().toISOString().split('T')[0],
                tecnico: utilizador,
                tempo_reparo: Math.round(tempoReparo * 10) / 10,
                satisfacao: Math.floor(Math.random() * 2) + 4, // NSS (a API não tem, mantemos simulado)
                sucesso: status === 'concluido' ? Math.random() > 0.1 : true,
                checkpoint: checkpoint,
                pendente_peca: pendentePeca === 'Sim' || pendentePeca === 'S',
                marca: getValor('marca') || '',
                modelo: getValor('modelo') || ''
            };
        });
        
        console.log(`✅ Processados ${dadosBrutos.length} registos`);
        console.log('📊 Distribuição por área:', contarPorArea(dadosBrutos));
        
        ultimaAtualizacao = new Date();
        atualizarInterface();
        
    } catch (erro) {
        console.error('Erro ao processar XML:', erro);
        carregarDadosExemplo();
    }
}

// Função auxiliar para contar registos por área
function contarPorArea(dados) {
    const contagem = {};
    dados.forEach(d => {
        contagem[d.area] = (contagem[d.area] || 0) + 1;
    });
    return contagem;
}

// Função para calcular KPIs (adaptada para usar os dados reais)
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

// Função para atualizar a interface (igual à anterior)
function atualizarInterface() {
    const kpis = calcularKPIs();
    
    // KPIs Globais
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
    
    // Atualiza totais por área
    document.getElementById('totalMobileCliente').textContent = kpis.mobileCliente?.total || 0;
    document.getElementById('totalMobileDG').textContent = kpis.mobileDG?.total || 0;
    document.getElementById('totalInformatica').textContent = kpis.informatica?.total || 0;
    document.getElementById('totalPequenosDomesticos').textContent = kpis.pequenosDomesticos?.total || 0;
    document.getElementById('totalSomImagem').textContent = kpis.somImagem?.total || 0;
    document.getElementById('totalEntretenimento').textContent = kpis.entretenimento?.total || 0;
    
    // Atualiza cada negócio
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
    
    // Rodapé
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

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Inicializando dashboard com campos reais...');
    
    document.getElementById('btnAtualizar').addEventListener('click', () => {
        document.getElementById('btnAtualizar').innerHTML = '<span class="material-icons">refresh</span> A atualizar...';
        buscarDadosAPI().finally(() => {
            document.getElementById('btnAtualizar').innerHTML = '<span class="material-icons">refresh</span> Atualizar';
        });
    });
    
    document.getElementById('periodoSelect').addEventListener('change', (e) => {
        console.log('Período alterado:', e.target.value);
        buscarDadosAPI();
    });
    
    buscarDadosAPI();
});
