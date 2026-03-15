// CONFIGURAÇÕES DA API
const API_URL = 'https://reportingwss.noshape.com/ServiceV3.asmx/TSC_OrsAbertasEmCadaCheckpoint_2';
const USERNAME = 'francisco.moreira@worten.pt';
const PASSWORD = 'Alice311020***';

// Estado da aplicação
let dadosBrutos = [];
let ultimaAtualizacao = null;

// Função para buscar dados da API
async function buscarDadosAPI() {
    try {
        console.log('🔄 Buscando dados da API...');
        
        // Prepara os parâmetros
        const parametros = new URLSearchParams();
        parametros.append('UserName', USERNAME);
        parametros.append('
