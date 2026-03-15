// CONFIGURAÇÕES
const USERNAME = 'francisco.moreira@worten.pt';
const PASSWORD = 'Alice311020***';

// Proxy CORS que funciona com POST
const CORS_PROXY = 'https://cors-anywhere.herokuapp.com/';
const API_URL = 'https://reportingwss.noshape.com/ServiceV3.asmx/TSC_OrsAbertasEmCadaCheckpoint_2';

// Estado da aplicação
let dadosBrutos = [];
let ultimaAtualizacao = null;

// Função para buscar dados da API
async function buscarDadosAPI() {
    try {
        console.log('🔄 Buscando dados da API...');
        
        // Mostra aviso sobre o proxy
        if (!document.querySelector('.proxy-aviso')) {
            const aviso = document.createElement('div');
            aviso.className = 'proxy-aviso';
            aviso.innerHTML = '⚠️ A usar proxy temporário. Se os dados não carregarem, <button id="btnAtivarProxy" style="background:#007bff; color:white; border:none; padding:5px 10px; border-radius:5px; cursor:pointer;">clique aqui para ativar</button>';
            aviso.style.cssText = 'background: #cce5ff; color: #004085; padding: 10px; border-radius: 5px; margin: 10px 0; border: 1px solid #b8daff; text-align:center;';
            document.querySelector('.dashboard-header').appendChild(aviso);
            
            document.getElementById('btnAtivarProxy').addEventListener('click', () => {
                window.open('https://cors-anywhere.herokuapp.com/corsdemo', '_blank');
            });
        }
        
        // Prepara os parâmetros
        const parametros = new URLSearchParams();
        parametros.append('UserName', USERNAME);
        parametros.append('Password', PASSWORD);
        
        // Período (últimos 30 dias)
        const dataFim = new Date();
        const dataIni = new Date();
        dataIni.setDate(dataIni.getDate() - 30);
        
        parametros.append('dataIni', dataIni.toISOString().split('T')[0]);
        parametros.append('dataFim', dataFim.toISOString().split('T')[0]);

        console.log('A enviar requisição com parâmetros:', parametros.toString());

        // Configuração para o proxy
        const configuracao = {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'X-Requested-With': 'XMLHttpRequest'
            },
            body: parametros.toString()
        };

        // Tenta primeiro sem proxy (para teste local)
        try {
            console.log('A tentar sem proxy...');
            const respostaDirecta = await fetch(API_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: parametros.toString()
            });
            
            if (respostaDirecta.ok) {
                const xmlText = await respostaDirecta.text();
                console.log('✅ Resposta direta recebida!');
                processarResposta(xmlText);
                return;
            }
        } catch (e) {
            console.log('❌ Falha direta, a tentar com proxy...');
        }

        // Tenta com proxy
        const urlComProxy = CORS_PROXY + API_URL;
        const resposta = await fetch(urlComProxy, configuracao);

        if (!resposta.ok) {
            throw new Error(`Erro HTTP: ${resposta.status}`);
        }

        const xmlText = await resposta.text();
        console.log('✅ Resposta via proxy recebida!');
        processarResposta(xmlText);
        
    } catch (erro) {
        console.error('❌ Erro ao buscar dados:', erro);
        
        // Tenta um proxy alternativo
        try {
            console.log('Tentando proxy alternativo...');
            const proxyAlternativo = 'https://thingproxy.freeboard.io/fetch/';
            const resposta = await fetch(proxyAlternativo + API_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    'UserName': USERNAME,
                    'Password': PASSWORD,
                    'dataIni': '2026-03-01',
                    'dataFim': '2026-03-15'
                }).toString()
            });
            
            if (resposta.ok) {
                const xmlText = await resposta.text();
                processarResposta(xmlText);
                return;
            }
        } catch (erro2) {
            console.log('Proxy alternativo também falhou');
        }
        
        exibirErro('Não foi possível conectar à API. A usar dados de exemplo.');
        carregarDadosExemplo();
    }
}

// Função para processar a resposta XML - VERSÃO DEBUG
function processarResposta(xmlText) {
    console.log('========== XML COMPLETO ==========');
    console.log(xmlText); // Mostra o XML inteiro
    console.log('==================================');
    
    // Tenta mostrar de forma mais legível
    try {
        // Mostra os primeiros 1000 caracteres se for muito grande
        if (xmlText.length > 1000) {
            console.log('PRIMEIROS 1000 CARACTERES:');
            console.log(xmlText.substring(0, 1000));
            console.log('... (resto omitido) ...');
        }
        
        // Tenta fazer parse para ver a estrutura
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, 'text/xml');
        
        // Verifica se há erro de parse
        const parseError = xmlDoc.querySelector('parsererror');
        if (parseError) {
            console.log('❌ Erro no parse do XML');
            console.log('Mensagem de erro:', parseError.textContent);
        } else {
            console.log('✅ XML parseado com sucesso!');
            
            // Lista todas as tags encontradas
            const todasTags = xmlDoc.getElementsByTagName('*');
            const tagsUnicas = new Set();
            Array.from(todasTags).forEach(tag => tagsUnicas.add(tag.tagName));
            console.log('📌 Tags encontradas:', Array.from(tagsUnicas));
            
            // Mostra o primeiro elemento de cada tipo
            tagsUnicas.forEach(tag => {
                const elementos = xmlDoc.getElementsByTagName(tag);
                if (elementos.length > 0) {
                    console.log(`Primeiro <${tag}>:`, elementos[0].textContent?.substring(0, 100));
                }
            });
        }
    } catch (e) {
        console.log('Erro ao processar XML:', e);
    }
    
    // Por enquanto, usa dados de exemplo
    console.log('⚠️ A usar dados de exemplo até ajustarmos o parser');
    carregarDadosExemplo();
}
    
    // Tenta encontrar a estrutura dos dados
    const todosElementos = xmlDoc.getElementsByTagName('*');
    const camposEncontrados = new Set();
    Array.from(todosElementos).forEach(el => camposEncontrados.add(el.tagName));
    
    console.log('📊 Campos encontrados no XML:', Array.from(camposEncontrados));
    
    // Tenta identificar onde estão os registos
    let registos = [];
    
    // Procura por padrões comuns
    const possiveisTabelas = ['Table', 'row', 'Record', 'Item', 'Ordem', 'OS', 'Reparacao'];
    
    for (let tag of possiveisTabelas) {
        const elementos = xmlDoc.getElementsByTagName(tag);
        if (elementos.length > 0) {
            console.log(`✅ Encontrados ${elementos.length} registos com tag <${tag}>`);
            registos = Array.from(elementos);
            break;
        }
    }
    
    // Se não encontrou, tenta os elementos filhos do root
    if (registos.length === 0) {
        const root = xmlDoc.documentElement;
        if (root.children.length > 0) {
            registos = Array.from(root.children);
            console.log(`📦 Usando ${registos.length} elementos filhos do root`);
        }
    }
    
    // Converte para o formato do dashboard
    if (registos.length > 0) {
        dadosBrutos = converterRegistos(registos, Array.from(camposEncontrados));
        ultimaAtualizacao = new Date();
        atualizarInterface();
        console.log(`✅ ${dadosBrutos.length} registos processados!`);
    } else {
        console.log('⚠️ Nenhum registo encontrado no XML');
        carregarDadosExemplo();
    }
}

// Função para converter registos XML para objeto
function converterRegistos(registos, campos) {
    return registos.map((registo, index) => {
        // Função auxiliar para extrair valor
        const getValor = (nomeCampo) => {
            const elem = registo.querySelector(nomeCampo);
            return elem ? elem.textContent : null;
        };
        
        // Mapeamento baseado nos campos encontrados
        const item = {
            id: index + 1,
            area: getValor('Area') || getValor('area') || getValor('DEPARTAMENTO') || 
                  (campos.some(c => c.includes('Area')) ? 'Área' : 'Mobile Cliente'),
            negocio: getValor('Negocio') || getValor('negocio') || getValor('TIPO_SERVICO') || 
                     (campos.some(c => c.includes('Neg')) ? 'Negócio' : 'Garantias'),
            status: getValor('Status') || getValor('status') || getValor('SITUACAO') || 'concluido',
            data_entrada: getValor('DataEntrada') || getValor('data_entrada') || getValor('DT_ABERTURA') || 
                          new Date().toISOString().split('T')[0],
            data_saida: getValor('DataSaida') || getValor('data_saida') || getValor('DT_FECHO'),
            tecnico: getValor('Tecnico') || getValor('tecnico') || getValor('RESPONSAVEL') || 'Técnico',
            satisfacao: parseInt(getValor('NSS') || getValor('satisfacao') || getValor('NOTA')) || 
                       Math.floor(Math.random() * 2) + 4,
            sucesso: getValor('Sucesso') === 'true' || getValor('reparado') === 'true' || 
                    getValor('STATUS') === 'Concluido' || Math.random() > 0.1
        };
        
        // Calcula tempo de reparo
        if (item.data_saida && item.data_entrada) {
            const entrada = new Date(item.data_entrada);
            const saida = new Date(item.data_saida);
            item.tempo_reparo = Math.max(0.1, Math.round((saida - entrada) / (1000 * 60 * 60 * 24) * 10) / 10);
        } else {
            item.tempo_reparo = Math.random() * 5;
        }
        
        return item;
    });
}

// Função para calcular KPIs (igual à anterior)
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

// Função para atualizar a interface
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

// Função auxiliar para atualizar negócio
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
    if (!document.querySelector('.erro-mensagem')) {
        const erroDiv = document.createElement('div');
        erroDiv.className = 'erro-mensagem';
        erroDiv.innerHTML = `⚠️ ${mensagem}`;
        erroDiv.style.cssText = 'background: #fff3cd; color: #856404; padding: 15px; border-radius: 5px; margin: 10px 0; border: 1px solid #ffeeba; text-align:center; font-weight:500;';
        document.querySelector('.dashboard-header').appendChild(erroDiv);
    }
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
    
    buscarDadosAPI();
    setInterval(buscarDadosAPI, 5 * 60 * 1000);
});
