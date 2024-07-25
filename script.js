document.getElementById('generateButton').addEventListener('click', generateNewExcel);

let excelData = [];
let jsonData = {};

document.getElementById('excelInput').addEventListener('change', handleExcelFile);
document.getElementById('jsonInput').addEventListener('change', handleJsonFile);

function handleExcelFile(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        excelData = XLSX.utils.sheet_to_json(worksheet);
    };
    reader.readAsArrayBuffer(file);
}

function handleJsonFile(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        jsonData = JSON.parse(e.target.result);
    };
    reader.readAsText(file);
}

function cleanString(str) {
    return str
        .replace(/[^a-zA-Z\s]/g, '')    // Remove caracteres não alfabéticos
        .replace(/\s+/g, ' ')           // Remove múltiplos espaços
        .trim()
        .toLowerCase();
}

function findClientData(clientName) {
    const cleanedClientName = cleanString(clientName);
    return Object.values(jsonData).find(item => {
        const cleanedJsonName = cleanString(item.dest.xNome);
        return cleanedJsonName.includes(cleanedClientName);
    });
}

function formatDateBR(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Janeiro é 0
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function generateNewExcel() {
    if (excelData.length === 0 || Object.keys(jsonData).length === 0) {
        alert('Por favor, insira a planilha Excel e o arquivo JSON.');
        return;
    }

    const today = new Date();
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    const newSheetData = excelData.map(row => {
        const clientData = findClientData(row.Cliente);

        return {
            Numero: row.Pedido,
            OrdemCarregamento: '',
            IdEmbarcador: 25574,
            IdRemetente: '',
            IdUnidade: 25700,
            EstimativaEntrega: formatDateBR(tomorrow),
            CnpjRemetente: clientData ? clientData.emit.CNPJ : '',
            DataPedido: formatDateBR(today),
            CnpjUnidade: '7216100000107',
            CnpjCliente: clientData ? clientData.dest.CNPJ : '',
            NomeCliente: row.Cliente,
            EnderecoCliente: clientData ? `${clientData.dest.xLgr}, ${clientData.dest.nro}` : '',
            CepCliente: clientData ? clientData.dest.cMun : '',
            BairroCliente: clientData ? clientData.dest.xBairro : '',
            CidadeCliente: clientData ? clientData.dest.xMun : '',
            EstadoCliente: clientData ? clientData.dest.UF : '',
            LatitudeCliente: '',
            LongitudeCliente: '',
            ValorPedido: '',
            PesoLiquido: row['Peso Bruto Estimado'],
            PesoBruto: row['Peso Bruto Estimado'],
            Cubagem: '',
            QtdeVolumes: '',
            QtdeCaixas: '',
            Observações: row.Pedido,
            Reentrega: 0,
            DataAgendamento: '',
            'Item-codigo': '',
            'Item-descricao': '',
            'Item-unidade-medida': '',
            'Item-quantidade': '',
            'Item-cubagem': '',
            'Item-peso_liquido': '',
            'Item-peso_bruto': '',
            'Item-valor-unitário': ''
        };
    });

    const worksheet = XLSX.utils.json_to_sheet(newSheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Planilha Gerada');
    XLSX.writeFile(workbook, 'PlanilhaFormatadaRavex.xlsx');

    const summary = document.getElementById('summary');
    summary.innerHTML = `
        <p>Total de registros processados: ${newSheetData.length}</p>
    `;
}
