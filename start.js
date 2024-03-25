var http = require('http');
//Required package
var pdf = require("pdf-creator-node")
var fs = require('fs')
const ExcelJS = require('exceljs');
  
// Read HTML Template
var html = fs.readFileSync('template.html', 'utf8')

const excelFilePath = 'IdentikitFinanziario.xlsx';
const workbook = new ExcelJS.Workbook();

var questions = [];
var responses = [];

var document = {
    html: html,
    data: {
    },
    path: ""
};

var options = { format: "A4", orientation: "portrait", border: "10mm" };


http.createServer(function (req, res) {
    res.writeHead(200, {'Content-Type': 'text/html'});
    res.end('Hello World!');

    // Read the Excel file
    workbook.xlsx.readFile(excelFilePath)
    .then(() => {
        // Access the worksheets, rows, and cells
        const worksheet = workbook.getWorksheet(1); // Assuming the first worksheet
        worksheet.eachRow((row, rowNumber) => {
            if(rowNumber == 1)// La prima riga contiene le domande
                row.eachCell((cell, cellNumber) => {
                    questions.push(cell.value);
                })
            //scorro le risposte
            else{
                responses = [];
                row.eachCell((cell, cellNumber) => {
                    responses.push(cell.value);
                });
                document.data = {
                    questions: questions,
                    responses: responses
                }
                document.path = "./IdentikitFinanziario_" + responses[4].replace(" ", "_") + ".pdf"
                pdf.create(document, options)
                    .then(res => {
                        console.log(res)
                    })
                    .catch(error => {
                        console.error(error)
                    });
                        }
                    });
    })
    .catch((error) => {
        console.error('Error reading Excel file:', error.message);
    });

    

}).listen(8080);