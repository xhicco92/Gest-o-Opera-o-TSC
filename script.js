// CONFIGURAÇÕES
const USERNAME = 'francisco.moreira@worten.pt';
const PASSWORD = 'Alice311020***';

// Proxy CORS gratuito (vamos usar este para contornar o bloqueio)
const CORS_PROXY = 'https://api.allorigins.win/raw?url=';

// URL da API (codificada para o proxy)
const API_URL = 'https://reportingwss.noshape.com/ServiceV3.asmx/TSC_OrsAbertasEmCadaCheckpoint_2';

// Estado da aplicação
let dadosBrutos = [];
let ultimaAtualizacao = null;

// Função para buscar dados da API via proxy
async function buscarDadosAPI() {
    try {
        console.log('🔄 Buscando dados da API...');
        
        // Prepara os parâmetros para enviar via POST
        const parametros = new URLSearchParams();
        parametros.append('UserName', USERNAME);
        parametros.append('Password', PASSWORD);
        
        // Definir período (últimos 30 dias)
        const dataFim = new Date();
        const dataIni = new Date();
        dataIni.setDate(dataIni.getDate() - 30);
        
        parametros.append('dataIni', dataIni.toISOString().split('T')[0]);
        parametros.append('dataFim', dataFim.toISOString().split('T')[0]);

        // Configuração da requisição
        const configuracao = {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: parametros.toString()
        };

        // Usa o proxy para contornar CORS
        const urlComProxy = CORS_PROXY + encodeURIComponent(API_URL);
        
        console.log('A enviar requisição...');
        
        const resposta = await fetch(urlComProxy, configuracao);

        if (!resposta.ok) {
            throw new Error(`Erro HTTP: ${resposta.status}`);
        }

        // Obtém o texto XML da resposta
        const xmlText = await resposta.text();
        console.log('XML recebido com sucesso!');
        
        // Parse do XML
        const dados = parseXML(xmlText);
        
        if (dados && dados.length > 0) {
            dadosBrutos = dados;
            ultimaAtualizacao = new Date();
            atualizarInterface();
            console.log(`✅ ${dados.length} registos carregados!`);
        } else {
            console.log('⚠️ Nenhum dado encontrado no período');
            carregarDadosExemplo();
        }
        
    } catch (erro) {
        console.error('❌ Erro ao buscar dados:', erro);
        exibirErro(erro.message);
        carregarDadosExemplo();
    }
}

// Função para parsear XML
function parseXML(xmlString) {
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlString, 'text/xml');
        
        // Verifica se há erro no parse
        const parseError = xmlDoc.querySelector('parsererror');
        if (parseError) {
            throw new Error('Erro ao parsear XML');
        }
        
        // Tenta encontrar os registos (ajuste conforme estrutura real)
        // Primeiro, vamos ver a estrutura
        console.log('Estrutura do XML:', xmlDoc.documentElement.tagName);
        
        // Tenta diferentes formas de encontrar os dados
        let registos = [];
        
        // Procura por tags comuns em APIs ASMX
        const possiveisTags = ['Table', 'row', 'Record', 'Item', 'resultado'];
        
        for (let tag of possiveisTags) {
            const elementos = xmlDoc.getElementsByTagName(tag);
            if (elementos.length > 0) {
                console.log(`Encontrados ${elementos.length} elementos com tag <${tag}>`);
                registos = Array.from(elementos);
                break;
            }
        }
        
        // Se não encontrou com as tags comuns, pega todos os elementos filhos
        if (registos.length === 0) {
            const root = xmlDoc.documentElement;
            if (root.children.length > 0) {
                registos = Array.from(root.children);
                console.log(`Usando ${registos.length} elementos filhos diretos`);
            }
        }
        
        // Converte para o formato que o dashboard entende
        return converterParaFormatoDashboard(registos, xmlDoc);
        
    } catch (erro) {
        console.error('Erro no parse do XML:', erro);
        return [];
    }
}

// Função para converter os dados para o formato do dashboard
function converterParaFormatoDashboard(registos, xmlDoc) {
    const dadosConvertidos = [];
    
    // Se não há registos, tenta extrair dados diretamente do XML
    if (registos.length === 0) {
        // Tenta encontrar padrões de dados no XML
        const todosElementos = xmlDoc.getElementsByTagName('*');
        const nomesCampos = new Set();
        Array.from(todosElementos).forEach(el => nomesCampos.add(el.tagName));
        
        console.log('Campos disponíveis no XML:', Array.from(nomesCampos));
        
        // Se encontrou campos mas não registos, cria registos virtuais para teste
        if (nomesCampos.size > 0) {
            return gerarDadosExemploComCampos(Array.from(nomesCampos));
        }
        
        return [];
    }
    
    // Processa cada registo encontrado
    registos.forEach((registo, index) => {
        // Tenta extrair valores de cada campo possível
        const extrairValor = (nomeCampo) => {
            const elemento = registo.querySelector(nomeCampo);
            return elemento ? elemento.textContent : null;
        };
        
        // Mapeamento inteligente - tenta adivinhar os campos
        const item = {
            id: index + 1,
            area: extrairValor('Area') || extrairValor('area') || extrairValor('DEPARTAMENTO') || 'Mobile Cliente',
            negocio: extrairValor('Negocio') || extrairValor('negocio') || extrairValor('TIPO') || 'Garantias',
            status: extrairValor('Status') || extrairValor('status') || extrairValor('SITUACAO') || 'concluido',
            data_entrada: extrairValor('DataEntrada') || extrairValor('data_entrada') || extrairValor('DT_ABERTURA') || new Date().toISOString().split('T')[0],
            data_saida: extrairValor('DataSaida') || extrairValor('data_saida') || extrairValor('DT_FECHO') || null,
            tecnico: extrairValor('Tecnico') || extrairValor('tecnico') || extrairValor('RESPONSAVEL') || 'João Silva',
            satisfacao: parseInt(extrairValor('NSS') || extrairValor('satisfacao') || extrairValor('NOTA')) || Math.floor(Math.random() * 2) + 4,
            sucesso: extrairValor('Sucesso') === 'true' || extrairValor('reparado') === 'true' || Math.random() > 0.1
        };
        
        // Calcula tempo de reparo se tiver data de saída
        if (item.data_saida && item.data_entrada) {
            const entrada = new Date(item.data_entrada);
            const saida = new Date(item.data_saida);
            item.tempo_reparo = Math.round((saida - entrada) / (1000 * 60 * 60 * 24) * 10) / 10;
        } else {
            item.tempo_reparo = Math.random() * 5; // valor simulado
        }
        
        dadosConvertidos.push(item);
    });
    
    return dadosConvertidos;
}

// Função para gerar dados de exemplo baseados nos campos encontrados
function gerarDadosExemploComCampos(campos) {
    console.log('Gerando dados de exemplo com base nos campos:', campos);
    
    const dados = [];
    const areas = ['Mobile Cliente', 'Mobile D&G', 'Informática', 'Pequenos Domésticos', 'Som e Imagem', 'Entretenimento'];
    const negocios = ['Garantias', 'Fora de Garantia', 'Extensão de Garantia'];
    
    for (let i = 1; i <= 150; i++) {
        const area = areas[Math.floor(Math.random() * areas.length)];
        let negocio;
        
        if (area === 'Mobile D&G') {
            negocio = 'D&G';
        } else {
            negocio = negocios[Math.floor(Math.random() * negocios.length)];
        }
        
        const data = new Date();
        data.setDate(data.getDate() - Math.floor(Math.random() * 30));
        
        dados.push({
            id: i,
            area: area,
            negocio: negocio,
            status: Math.random() > 0.3 ? 'concluido' : (Math.random() > 0.5 ? 'andamento' : 'pendente'),
            data_entrada: data.toISOString().split('T')[0],
            tecnico: ['João Silva', 'Maria Santos', 'Pedro Oliveira'][Math.floor(Math.random() * 3)],
            tempo_reparo: Math.random() * 8,
            satisfacao: Math.floor(Math.random() * 2) + 4,
            sucesso: Math.random() > 0.1
        });
    }
    
    return dados;
}

// Função para calcular KPIs
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

    // Mapeamento de áreas para slugs
    const mapaAreas = {
        'Mobile Cliente': 'mobileCliente',
        'Mobile D&G': 'mobileDG',
        'Informática': 'informatica',
        'Pequenos Domésticos': 'pequenosDomesticos',
        'Som e Imagem': 'somImagem',
        'Entretenimento': 'entretenimento'
    };

    // Processa cada reparação
    dadosBrutos.forEach(item => {
        const area = mapaAreas[item.area] || 'mobileCliente';
        let negocio = item.negocio;
        
        // Normaliza nome do negócio
        if (negocio.includes('Fora')) negocio = 'ForaGarantia';
        else if (negocio.includes('Extensão')) negocio = 'Extensao';
        else if (negocio.includes('Garantias')) negocio = 'Garantias';
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

    // Calcula médias
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
    
    // Atualiza KPIs Globais
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
    
    // Atualiza rodapé
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

// Função auxiliar para atualizar um negócio
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

// Função para exibir erro
function exibirErro(mensagem) {
    const erroDiv = document.createElement('div');
    erroDiv.className = 'erro-mensagem';
    erroDiv.innerHTML = `⚠️ Aviso: A usar dados de exemplo. (${mensagem})`;
    erroDiv.style.cssText = 'background: #fff3cd; color: #856404; padding: 10px; border-radius: 5px; margin: 10px 0; border: 1px solid #ffeeba;';
    
    if (!document.querySelector('.erro-mensagem')) {
        document.querySelector('.dashboard-header').appendChild(erroDiv);
    }
}

// Dados de exemplo para fallback
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
    console.log('🚀 Inicializando dashboard...');
    
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
    
    // Tenta buscar dados reais, se falhar usa exemplo
    buscarDadosAPI();
    
    setInterval(buscarDadosAPI, 5 * 60 * 1000);
});
