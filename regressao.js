const XLSX = require('xlsx');
 
/**
 * Função para ler as colunas de um arquivo Excel e retorná-las como vetores x e y.
 * @param {string} caminhoArquivo - O caminho para o arquivo .xlsx (ex: './dados.xlsx').
 * @param {string} nomeColunaX - O nome exato da coluna para os dados de X.
 * @param {string} nomeColunaY - O nome exato da coluna para os dados de Y.
 * @returns {{x: number[], y: number[]}} Um objeto contendo os vetores x e y.
 */
function lerDadosExcel(caminhoArquivo, nomeColunaX, nomeColunaY) {
    try {
        const workbook = XLSX.readFile(caminhoArquivo);
 
        const nomePlanilha = workbook.SheetNames[0];
        const planilha = workbook.Sheets[nomePlanilha];
 
        const dadosJson = XLSX.utils.sheet_to_json(planilha);
 
        const x = dadosJson.map(linha => linha[nomeColunaX]);
        const y = dadosJson.map(linha => linha[nomeColunaY]);
 
        if (x.some(v => v === undefined) || y.some(v => v === undefined)) {
            throw new Error(`Verifique se os nomes das colunas "${nomeColunaX}" e "${nomeColunaY}" estão corretos e existem no arquivo.`);
        }
 
        return { x, y };
    } catch (error) {
        console.error("Ocorreu um erro ao ler o arquivo Excel:", error.message);
        return { x: [], y: [] };
    }
}
 
function regressaoLinear(x, y) {
    if (x.length !== y.length || x.length === 0) {
        throw new Error("Os vetores x e y devem ter o mesmo tamanho e não podem ser vazios.");
    }
 
    const n = x.length;
    const media = arr => arr.reduce((a, b) => a + b, 0) / arr.length;
    const mediaX = media(x);
    const mediaY = media(y);
 
    let numerador = 0;
    let denominador = 0;
 
    for (let i = 0; i < n; i++) {
        numerador += (x[i] - mediaX) * (y[i] - mediaY);
        denominador += (x[i] - mediaX) ** 2;
    }
 
    const a = numerador / denominador;
    const b = mediaY - a * mediaX;
 
    let somaTotal = 0;
    let somaResiduos = 0;
 
    for (let i = 0; i < n; i++) {
        const yEstimado = a * x[i] + b;
        somaTotal += (y[i] - mediaY) ** 2;
        somaResiduos += (y[i] - yEstimado) ** 2;
    }
 
    const r2 = 1 - (somaResiduos / somaTotal);
 
    return {
        equacao: `y = ${a.toFixed(4)}x + ${b.toFixed(4)}`,
        a: a,
        b: b,
        r2: r2
    };
}
 
const caminhoDoArquivo = './analise_peso_inicial_final.xlsx'; 
const colunaX = 'Peso Inicial (kg)';
const colunaY = 'Peso Final (kg)';
 
const { x, y } = lerDadosExcel(caminhoDoArquivo, colunaX, colunaY);
 

if (x.length > 0 && y.length > 0) {
    
    const resultado = regressaoLinear(x, y);
 
    console.log("Análise de Regressão Linear");
    console.log("----------------------------");
    console.log(`Dados carregados do arquivo: ${caminhoDoArquivo}`);
    console.log(`Total de registros: ${x.length}`);
    console.log("Equação da reta:", resultado.equacao);
    console.log("Coeficiente angular (a):", resultado.a.toFixed(4));
    console.log("Coeficiente linear (b):", resultado.b.toFixed(4));
    console.log("Coeficiente de Determinação (R²):", resultado.r2.toFixed(4));
} else {
    console.log("A análise não pôde ser executada pois os dados não foram carregados.");
}