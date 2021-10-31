// npm init 
// npm install pdf-lib
// npm install xlsx
// npm install nodemailer
// node mainDEMO.js 
// instead of main.js uploading mainDemo.js difference is Credentials

let pdf = require("pdf-lib");
let xlsx = require("xlsx");
let fs = require("fs");
let nodemailer = require("nodemailer");
const { google } = require('googleapis');
const CLIENT_ID = 'yourClientID';
const CLIENT_SECRET = 'yourClientSecret';
const REDIRECT_URI = 'https://developers.google.com/oauthplayground';
const REFRESH_TOKEN = 'yourRefreshToken';

const oAuth2Client = new google.auth.OAuth2(CLIENT_ID,CLIENT_SECRET, REDIRECT_URI);
oAuth2Client.setCredentials({ refresh_token: REFRESH_TOKEN });

// READING EXCEL DATA GATHER VIA GOOGLE FORM
let dataXlsx = xlsx.readFile("data.xlsx");
let dataSheet = dataXlsx.Sheets['Form Responses 1'];
let data = xlsx.utils.sheet_to_json(dataSheet);

// READING DATE, SUBJECT AND CONTENT FOR CERTIFICATE
let date = "24 October 2021";
let subject = fs.readFileSync('subject.txt', 'utf-8');
let body = fs.readFileSync('body.txt', 'utf-8');    


for(let i = 0; i< data.length; i++){
    let name = data[i].Name;
    // CREATING & STORING CERTIFICATE
    createCertificate(name,date);

    let certiPath = "Certificates/"+ name + ".pdf";
    // READING EMAIL FROM JSON FOR ith USER
    let userEmail = data[i].Email;

    // CALLING FUNCTION PROMISE FOR SENDING MAIL
    sendMail(subject,body, name, userEmail, certiPath).then(function(result){
        console.log("Email is sent", result);
    }).catch(function(err){
        console.log(err);
    })
}

// FUNCTION FOR CREATING CERTIFICATE
function createCertificate(name, date){
    let pdfBytes = fs.readFileSync("template.pdf");
    let pdfPromise = pdf.PDFDocument.load(pdfBytes);
    pdfPromise.then(function(pdfdoc){
        let page = pdfdoc.getPage(0);
        page.drawText(name,{
            x: 105,
            y: 300,
            size: 50 
        });
        page.drawText(date,{
            x: 200,
            y: 218,
            size: 12 
        });

        let finalPdf = pdfdoc.save();
        let pdfName = "Certificates/"+name + ".pdf";
        finalPdf.then(function(finalBytes){
            fs.writeFileSync(pdfName,finalBytes);
        })
    })
}

// FUNCTION FOR SENDING EMAIL
async function sendMail(subject,body, name, userEmail, certiPath) {

    try {
        let accessToken = await oAuth2Client.getAccessToken();

        const transport = nodemailer.createTransport({
            service: 'gmail',
            auth: {
                type: 'OAuth2',
                user: 'udyanojha@gmail.com',
                clientId: CLIENT_ID,
                clientSecret: CLIENT_SECRET,
                refreshToken: REFRESH_TOKEN,
                accessToken : accessToken
            }
        })

        let mailOptions = {
            from: 'Udyan Ojha <udyanojha@gmail.com>',
            to: userEmail,
            subject: subject,
            html: "Dear "+name+"\n"+body,
            attachments: [{
                filename: name+".pdf",
                path: certiPath,
                contentType: 'application/pdf'
            }]
        }

        let result = await transport.sendMail(mailOptions);
        return result;

    } catch(err) {
        console.log(err);
        return err;
    }

}