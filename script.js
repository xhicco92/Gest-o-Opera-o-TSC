// CONFIGURAÇÕES
const USERNAME = 'francisco.moreira@worten.pt';
const PASSWORD = 'Alice311020***';

// Estado da aplicação
let dadosBrutos = [];
let ultimaAtualizacao = null;

// Função para testar diferentes formatos de data
async function buscarDadosAPI() {
    try {
        console.log('🔄 Testando diferentes formatos de data...');
        
        const dataFim = new Date();
        const dataIni = new Date();
        dataIni.setDate(dataIni.getDate() - 30);
        
        // Diferentes formatos de data para testar
        const formatosData = [
            { nome: 'ISO (aaaa-mm-dd)', ini: dataIni.toISOString().split('T')[0], fim: dataFim.toISOString().split('T')[0] },
            { nome: 'PT (dd/mm/aaaa)', ini: formatarDataPT(dataIni), fim: formatarDataPT(dataFim) },
            { nome: 'US (mm/dd/aaaa)', ini: formatarDataUS(dataIni), fim: formatarDataUS(dataFim) },
            { nome: 'Sem separador (aaaammdd)', ini: formatarDataNumero(dataIni), fim: formatarDataNumero(dataFim) }
        ];
        
        // Testa cada formato
        for (let formato of formatosData) {
            console.log(`\n📅 Testando formato: ${formato.nome}`);
            console.log(`dataIni: ${formato.ini}, dataFim: ${formato.fim}`);
            
            const parametros = new URLSearchParams();
            parametros.append('UserName', USERNAME);
            parametros.append('Password', PASSWORD);
            parametros.append('dataIni', formato.ini);
            parametros.append('dataFim', formato.fim);
            
            try {
                // Tenta sem proxy primeiro
                const resposta = await fetch('https://reportingwss.noshape.com/ServiceV3.asmx/TSC_OrsAbertasEmCadaCheckpoint_2', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                    },
                    body: parametros.toString()
                });
                
                console.log(`Status: ${resposta.status} ${resposta.statusText}`);
                
                if (resposta.ok) {
                    const texto = await resposta.text();
                    console.log('✅ SUCESSO! Resposta recebida:');
                    console.log(texto.substring(0, 500));
                    
                    // Se conseguir, processa os dados
                    processarResposta(texto);
                    return;
                } else {
                    console.log(`❌ Falhou com status ${resposta.status}`);
                    // Tenta ler o corpo do erro
                    try {
                        const erroTexto = await resposta.text();
                        console.log('Corpo do erro:', erroTexto.substring(0, 200));
                    } catch (e) {}
                }
            } catch (erro) {
                console.log(`❌ Erro de rede: ${erro.message}`);
            }
        }
        
        console.log('\n⚠️ Todos os formatos falharam. A usar dados de exemplo.');
        carregarDadosExemplo();
        
    } catch (erro) {
        console.error('Erro geral:', erro);
        carregarDadosExemplo();
    }
}

// Funções auxiliares para formatar datas
function formatarDataPT(data) {
    const dia = data.getDate().toString().padStart(2, '0');
    const mes = (data.getMonth() + 1).toString().padStart(2, '0');
    const ano = data.getFullYear();
    return `${dia}/${mes}/${ano}`;
}

function formatarDataUS(data) {
    const mes = (data.getMonth() + 1).toString().padStart(2, '0');
    const dia = data.getDate().toString().padStart(2, '0');
    const ano = data.getFullYear();
    return `${mes}/${dia}/${ano}`;
}

function formatarDataNumero(data) {
    const ano = data.getFullYear();
    const mes = (data.getMonth() + 1).toString().padStart(2, '0');
    const dia = data.getDate().toString().padStart(2, '0');
    return `${ano}${mes}${dia}`;
}

// Função para processar resposta (versão simplificada)
function processarResposta(xmlText) {
    console.log('\n========== RESPOSTA DA API ==========');
    console.log(xmlText.substring(0, 1000));
    console.log('======================================\n');
    
    // Mesmo com sucesso, por enquanto usamos dados de exemplo
    // até vermos a estrutura real
    console.log('⚠️ Por enquanto, a usar dados de exemplo');
    carregarDadosExemplo();
}

// Função para calcular KPIs (igual à anterior - mantém-se)
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
});
