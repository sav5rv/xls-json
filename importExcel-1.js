const XLSX = require("xlsx");
const { MongoClient } = require("mongodb");

// 📌 Configuração do MongoDB
const MONGO_URL = "mongodb://192.168.19.128:27017"; // Altere conforme necessário
const DATABASE_NAME = "meu_banco_de_dados";
const COLLECTION_NAME = "minha_colecao_1";

// 📌 Caminho do arquivo Excel
const FILE_PATH = "Acidentes.xls"; // Altere para o nome do seu arquivo

async function importExcelToMongoDB() {
    try {
        // 📌 Conectar ao MongoDB
        const client = new MongoClient(MONGO_URL);
        await client.connect();
        console.log("✅ Conectado ao MongoDB!");

        const db = client.db(DATABASE_NAME);
        const collection = db.collection(COLLECTION_NAME);

        // 📌 Ler o arquivo Excel
        const workbook = XLSX.readFile(FILE_PATH, { cellDates: true });
        const sheetName = workbook.SheetNames[0]; // Pega a primeira aba
        const sheet = workbook.Sheets[sheetName];

        // 📌 Converter os dados para JSON
        const data = XLSX.utils.sheet_to_json(sheet);
        console.log(data);

        // 📌 Inserir os dados no MongoDB
        if (data.length > 0) {
            await collection.insertMany(data);
            console.log(`✅ ${data.length} documentos inseridos com sucesso!`);
        } else {
            console.log("⚠️ Nenhum dado encontrado no arquivo.");
        }

        // 📌 Fechar a conexão com o banco
        await client.close();
        console.log("🔌 Conexão encerrada.");
    } catch (error) {
        console.error("❌ Erro ao importar dados:", error);
    }
}

// Executar a função de importação
importExcelToMongoDB();
