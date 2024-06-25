import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

interface ExcelRow {
    [key: string]: any;
}

function excelToJson(excelFilePath: string, outputDir: string): void {
    try {
        const fileBuffer = fs.readFileSync(excelFilePath);
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });

        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData: ExcelRow[] = XLSX.utils.sheet_to_json(worksheet, { defval: null });
            const jsonFileName = path.join(outputDir, `${sheetName}.json`);
            fs.writeFileSync(jsonFileName, JSON.stringify(jsonData, null, 4));
            console.log(`JSON gerado com sucesso para a aba '${sheetName}' em: ${jsonFileName}`);
        });
    } catch (error) {
        console.error('Erro ao converter o arquivo Excel para JSON:', error);
    }
}

const excelFilePath = './Margarida.xlsx'; // Ajuste o caminho conforme necessário
const outputDir = './json_output';        // Ajuste o diretório de saída conforme necessário
excelToJson(excelFilePath, outputDir);
