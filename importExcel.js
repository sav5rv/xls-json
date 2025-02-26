const XLSX = require("xlsx");
const { MongoClient } = require("mongodb");
const readline = require("readline");

// 📌 Criar interface para ler entrada do terminal
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

// 📌 Configuração do MongoDB
const MONGO_URL = "mongodb://192.168.19.128:27017"; // Altere conforme necessário
const DATABASE_NAME = "meu_banco_de_dados";
const COLLECTION_NAME = "minha_colecao_2";

// 📌 Perguntar ao usuário o nome do arquivo
rl.question("Digite o nome do arquivo XLS/XLSX (ex: dados.xlsx): ", async (filePath) => {
    try {
        // 📌 Conectar ao MongoDB
        const client = new MongoClient(MONGO_URL);
        await client.connect();
        console.log("✅ Conectado ao MongoDB!");

        const db = client.db(DATABASE_NAME);
        const collection = db.collection(COLLECTION_NAME);

        // 📌 Ler o arquivo Excel
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0]; // Pega a primeira aba
        const sheet = workbook.Sheets[sheetName];

        // 📌 Converter os dados para JSON
        const data = XLSX.utils.sheet_to_json(sheet);

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
    } finally {
        rl.close(); // Fechar o readline
    }
});
