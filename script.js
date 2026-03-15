// CONFIGURAÇÕES
const USERNAME = 'francisco.moreira@worten.pt';
const PASSWORD = 'Alice311020***';
const API_URL = 'https://reportingwss.noshape.com/ServiceV3.asmx/TSC_OrsAbertasEmCadaCheckpoint';
const PROXY_URL = 'https://cors-anywhere.herokuapp.com/';

// Estado
let dadosBrutos = [];
let ultimaAtualizacao = null;

// Função para buscar dados
async function buscarDadosAPI() {
    try {
        console.log('🔄 Buscando dados da API (formato Excel)...');
        
        // Calcular datas (últimos 30 dias)
        const dataFim = new Date();
        const dataIni = new Date();
        dataIni.setDate(dataIni.getDate() - 30);
        
        // Formatar datas como ISO (igual ao Excel)
        const dataIniStr = dataIni.toISOString().split('T')[0];
        const dataFimStr = dataFim.toISOString().split('T')[0];
        
        console.log(`📅 Período: ${dataIniStr} a ${dataFimStr}`);
        
        const corpo = new URLSearchParams({
            'UserName': USERNAME,
            'Password': PASSWORD,
            'dataIni': dataIniStr,
            'dataFim': dataFimStr
        });
        
        // Fazer pedido via proxy
        const resposta = await fetch(PROXY_URL + API_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Accept': 'application/vnd.ms-excel, application/xml, text/xml',
                'User-Agent': 'Microsoft Excel/16.0'
            },
            body: corpo.toString()
        });
        
        if (!resposta.ok) {
            throw new Error(`HTTP ${resposta.status}`);
        }
        
        // Obter resposta como texto (vai ser XML/Excel)
        const conteudo = await resposta.text();
        console.log('✅ Resposta recebida, tamanho:', conteudo.length, 'caracteres');
        
        // Verificar se é um erro conhecido
        if (conteudo.includes('Unauthorized')) {
            throw new Error('Acesso não autorizado - verifique permissões');
        }
        
        // Processar o conteúdo como Excel/XML
        processarConteudoExcel(conteudo);
        
    } catch (erro) {
        console.error('❌ Erro:', erro);
        
        // Mostrar aviso com instruções
        if (!document.querySelector('.proxy-aviso')) {
            const aviso = document.createElement('div');
            aviso.className = 'proxy-aviso';
            aviso.innerHTML = `
                <div style="background:#fff3cd; color:#856404; padding:15px; border-radius:5px; margin:10px 0; border:1px solid #ffeeba;">
                    ⚠️ Erro: ${erro.message}<br>
                    <strong>Instruções:</strong><br>
                    1. <a href="https://cors-anywhere.herokuapp.com/corsdemo" target="_blank">Clique aqui</a> e peça acesso temporário ao proxy<br>
                    2. Depois volte e clique em "Atualizar"<br>
                    3. Se o erro persistir, contacte o IT sobre permissões para o método TSC_OrsAbertasEmCadaCheckpoint
                </div>
            `;
            document.querySelector('.dashboard-header').appendChild(aviso);
        }
        
        carregarDadosExemplo();
    }
}

// Função para processar conteúdo Excel/XML
function processarConteudoExcel(conteudo) {
    try {
        // Parse do XML
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(conteudo, 'text/xml');
        
        // Verificar se é um ficheiro Excel (tem Workbook/Worksheet)
        const sheets = xmlDoc.getElementsByTagName('Worksheet');
        const rows = xmlDoc.getElementsByTagName('Row');
        
        console.log('📊 Estrutura do documento:');
        console.log('- Worksheets:', sheets.length);
        console.log('- Rows:', rows.length);
        
        // Procurar pelos dados - podem estar em diferentes formatos
        let dadosEncontrados = false;
        
        // Formato 1: Tabelas HTML/XML
        const tabelas = xmlDoc.getElementsByTagName('Table');
        if (tabelas.length > 0) {
            console.log('✅ Encontradas', tabelas.length, 'tabelas');
            dadosEncontrados = extrairDadosTabelas(tabelas, xmlDoc);
        }
        
        // Formato 2: Linhas diretamente
        if (!dadosEncontrados && rows.length > 0) {
            console.log('✅ Encontradas', rows.length, 'linhas');
            dadosEncontrados = extrairDadosLinhas(rows);
        }
        
        // Formato 3: Registos XML (Table do ADO.NET)
        if (!dadosEncontrados) {
            const registos = xmlDoc.getElementsByTagName('Table');
            if (registos.length > 0) {
                console.log('✅ Encontrados', registos.length, 'registos');
                dadosEncontrados = extrairDadosRegistos(registos);
            }
        }
        
        if (!dadosEncontrados) {
            throw new Error('Não foi possível extrair dados do formato recebido');
        }
        
        ultimaAtualizacao = new Date();
        atualizarInterface();
        console.log('✅ Dashboard atualizado com sucesso!');
        
    } catch (erro) {
        console.error('❌ Erro ao processar Excel/XML:', erro);
        console.log('Primeiros 500 caracteres do conteúdo:', conteudo.substring(0, 500));
        carregarDadosExemplo();
    }
}

// Função para extrair dados de tabelas
function extrairDadosTabelas(tabelas, xmlDoc) {
    try {
        // Procurar pela primeira tabela com dados
        for (let tabela of tabelas) {
            const linhas = tabela.getElementsByTagName('Row');
            if (linhas.length > 1) { // Tem cabeçalho + dados
                console.log(`📋 Tabela com ${linhas.length} linhas`);
                
                // Extrair cabeçalhos (primeira linha)
                const celulasCabecalho = linhas[0].getElementsByTagName('Cell');
                const cabecalhos = Array.from(celulasCabecalho).map(cell => {
                    const data = cell.getElementsByTagName('Data')[0];
                    return data ? data.textContent : '';
                });
                
                console.log('📌 Cabeçalhos:', cabecalhos);
                
                // Extrair dados (restantes linhas)
                dadosBrutos = [];
                for (let i = 1; i < linhas.length; i++) {
                    const celulas = linhas[i].getElementsByTagName('Cell');
                    const linha = {};
                    
                    Array.from(celulas).forEach((cell, idx) => {
                        if (idx < cabecalhos.length) {
                            const data = cell.getElementsByTagName('Data')[0];
                            linha[cabecalhos[idx]] = data ? data.textContent : '';
                        }
                    });
                    
                    // Converter para formato do dashboard
                    const item = converterLinhaParaFormato(linha, i);
                    if (item) dadosBrutos.push(item);
                }
                
                return dadosBrutos.length > 0;
            }
        }
    } catch (e) {
        console.log('Erro ao extrair tabelas:', e);
    }
    return false;
}

// Função para extrair dados de linhas
function extrairDadosLinhas(rows) {
    try {
        if (rows.length < 2) return false;
        
        // Primeira linha são cabeçalhos
        const celulasCabecalho = rows[0].getElementsByTagName('Cell');
        const cabecalhos = Array.from(celulasCabecalho).map(cell => {
            const data = cell.getElementsByTagName('Data')[0];
            return data ? data.textContent : '';
        });
        
        dadosBrutos = [];
        for (let i = 1; i < rows.length; i++) {
            const celulas = rows[i].getElementsByTagName('Cell');
            const linha = {};
            
            Array.from(celulas).forEach((cell, idx) => {
                if (idx < cabecalhos.length) {
                    const data = cell.getElementsByTagName('Data')[0];
                    linha[cabecalhos[idx]] = data ? data.textContent : '';
                }
            });
            
            const item = converterLinhaParaFormato(linha, i);
            if (item) dadosBrutos.push(item);
        }
        
        return dadosBrutos.length > 0;
    } catch (e) {
        console.log('Erro ao extrair linhas:', e);
        return false;
    }
}

// Função para extrair dados de registos XML
function extrairDadosRegistos(registos) {
    try {
        dadosBrutos = Array.from(registos).map((reg, idx) => {
            const linha = {};
            
            // Extrair todos os campos do registo
            Array.from(reg.children).forEach(child => {
                linha[child.tagName] = child.textContent;
            });
            
            return converterLinhaParaFormato(linha, idx);
        }).filter(item => item !== null);
        
        return dadosBrutos.length > 0;
    } catch (e) {
        console.log('Erro ao extrair registos:', e);
        return false;
    }
}

// Função para converter linha do Excel para formato do dashboard
function converterLinhaParaFormato(linha, index) {
    try {
        // Mapear campos com base no código M do Excel
        const area = linha['area'] || 'Mobile Cliente';
        const tipoGarantia = linha['tipo_garantia'] || 'Garantias';
        const dataCheckin = linha['data_checkin'];
        const checkpoint = linha['checkpoint_atual'] || '';
        const utilizador = linha['utilizador'] || 'Técnico';
        const entdade = linha['entdade'] || '';
        
        // Filtrar entidades (igual ao Excel: <> "Recondicionado PT" and <> "TSC")
        if (entdade === 'Recondicionado PT' || entdade === 'TSC') {
            return null;
        }
        
        // Determinar status
        let status = 'pendente';
        if (checkpoint) {
            if (checkpoint.includes('FECHADO') || checkpoint.includes('CONCLUIDO') || checkpoint.includes('FINALIZADO')) {
                status = 'concluido';
            } else if (checkpoint.includes('REPARACAO') || checkpoint.includes('ANALISE') || checkpoint.includes('TRIAGEM')) {
                status = 'andamento';
            }
        }
        
        // Calcular TAT (igual ao Excel)
        let tempoReparo = 0;
        if (dataCheckin) {
            const dataEntrada = new Date(dataCheckin);
            const hoje = new Date();
            const diffTime = Math.abs(hoje - dataEntrada);
            tempoReparo = diffTime / (1000 * 60 * 60 * 24); // dias
        }
        
        // Normalizar área
        let areaNorm = area;
        if (area.includes('Mobile') || area.includes('TELEMOVEL')) areaNorm = 'Mobile Cliente';
        else if (area.includes('D&G') || area.includes('DG')) areaNorm = 'Mobile D&G';
        else if (area.includes('Informatica') || area.includes('PC') || area.includes('NOTEBOOK')) areaNorm = 'Informática';
        else if (area.includes('Domestico') || area.includes('PDA') || area.includes('ELECTRODOMESTICO')) areaNorm = 'Pequenos Domésticos';
        else if (area.includes('Som') || area.includes('Imagem') || area.includes('TV') || area.includes('AUDIO')) areaNorm = 'Som e Imagem';
        else if (area.includes('Entretenimento') || area.includes('GAMING') || area.includes('CONSOLA')) areaNorm = 'Entretenimento';
        
        // Normalizar negócio
        let negocioNorm = 'Garantias';
        if (tipoGarantia) {
            if (tipoGarantia.includes('Fora') || tipoGarantia.includes('FORA')) negocioNorm = 'Fora de Garantia';
            else if (tipoGarantia.includes('Extensao') || tipoGarantia.includes('EXTENSAO')) negocioNorm = 'Extensão de Garantia';
            else if (tipoGarantia.includes('Garantia') || tipoGarantia.includes('GARANTIA')) negocioNorm = 'Garantias';
        }
        
        // Ajustar para Mobile D&G
        if (areaNorm === 'Mobile D&G') {
            negocioNorm = 'D&G';
        }
        
        return {
            id: index + 1,
            area: areaNorm,
            negocio: negocioNorm,
            status: status,
            data_entrada: dataCheckin ? dataCheckin.split('T')[0] : new Date().toISOString().split('T')[0],
            tecnico: utilizador,
            tempo_reparo: Math.round(tempoReparo * 10) / 10,
            satisfacao: 4, // NSS não disponível na API
            sucesso: status === 'concluido',
            checkpoint: checkpoint,
            marca: linha['marca'] || '',
            modelo: linha['modelo'] || '',
            sn: linha['sn_equipamento'] || '',
            entidade: entdade
        };
    } catch (e) {
        console.log('Erro ao converter linha:', e);
        return null;
    }
}

// Funções de KPIs e interface (mantidas iguais às versões anteriores)
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

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Inicializando dashboard com parser Excel...');
    
    // Botão de atualização
    document.getElementById('btnAtualizar').addEventListener('click', () => {
        document.getElementById('btnAtualizar').innerHTML = '<span class="material-icons">refresh</span> A atualizar...';
        buscarDadosAPI().finally(() => {
            document.getElementById('btnAtualizar').innerHTML = '<span class="material-icons">refresh</span> Atualizar';
        });
    });
    
    // Seletor de período
    document.getElementById('periodoSelect').addEventListener('change', (e) => {
        console.log('Período alterado:', e.target.value);
        buscarDadosAPI();
    });
    
    // Iniciar
    buscarDadosAPI();
    
    // Atualização automática a cada 5 minutos
    setInterval(buscarDadosAPI, 5 * 60 * 1000);
});
