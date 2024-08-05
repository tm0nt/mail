const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Configurações do serviço SMTP
const transporter = nodemailer.createTransport({
    host: 'smtp.viasmtp.com.br',
    port: 587,
    secure: false, // true para 465, false para outras portas
    auth: {
        user: 'marcio15515',
        pass: 'flamengo10'
    }
});

// Função para enviar email
async function sendEmail(to, subject, htmlContent) {
    const mailOptions = {
        from: '"Transportadora Correios" <transporte@alfandesitegabr.>',
        to: to,
        subject: subject,
        html: htmlContent
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log(`Email enviado para: ${to}`);
    } catch (error) {
        console.error(`Erro ao enviar email para ${to}:`, error);
    }
}

// Função para ler a planilha e enviar emails
async function sendEmailsFromExcelFile(filePath, htmlFilePath) {
    // Ler o conteúdo HTML do arquivo
    const htmlTemplate = fs.readFileSync(htmlFilePath, 'utf-8');

    // Ler a planilha Excel
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);

    const batchSize = 1;
    let batch = [];

    for (let i = 0; i < data.length; i++) {
        const { Email: email, Nome: nome } = data[i];

        if (email) { // Verificar se o email existe
            // Substituir o marcador de posição no HTML pelo nome real
            const personalizedHtml = htmlTemplate.replace('${nome}', nome);

            batch.push({ email, personalizedHtml });

            if (batch.length === batchSize || i === data.length - 1) {
                const to = batch.map(item => item.email).join(', ');

                await sendEmail(
                    to,
                    'Urgente: Seu Pedido Foi Retido',
                    batch.map(item => item.personalizedHtml).join('<br><br>')
                );
                batch = [];
            }
        }
    }
}

// Caminho para a planilha Excel
const excelFilePath = './leads2.xlsx';
// Caminho para o arquivo HTML
const htmlFilePath = path.join(__dirname, 'email_template.html');

// Enviar emails
sendEmailsFromExcelFile(excelFilePath, htmlFilePath).then(() => {
    console.log('Todos os emails foram enviados.');
}).catch((error) => {
    console.error('Erro ao enviar emails:', error);
});
