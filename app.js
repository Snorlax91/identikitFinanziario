var express = require('express');
var app = express()
app.set('port', 8080);
//Required package
var pdf = require("pdf-creator-node")
var fs = require('fs')
const archiver = require('archiver');
const ExcelJS = require('exceljs');
  
// Read HTML Template
var html = fs.readFileSync('template.html', 'utf8')

const excelFilePath = 'IdentikitFinanziario.xlsx';
const workbook = new ExcelJS.Workbook();
const dirname = "./identikit/";

var questions = [];
var responses = [];

var document = {
    html: html,
    data: {
    },
    path: ""
};

var options = { format: "A4", orientation: "portrait", border: "10mm" };

app.get('/download', function (req, res) {
    var filePath = "/my/file/path/..."; // Or format the path using the `id` rest param
    var fileName = "report.pdf"; // The default name the browser will use
    var filesName = [];
    var workedRows = 0;
    var rowsCount = 99;
    // Read the Excel file
    workbook.xlsx.readFile(excelFilePath)
    .then(() => {
        // Access the worksheets, rows, and cells
        const worksheet = workbook.getWorksheet(1); // Assuming the first worksheet
        rowsCount = worksheet.rowCount;
        worksheet.eachRow((row, rowNumber) => {
            if(rowNumber == 1){// La prima riga contiene le domande
                row.eachCell((cell, cellNumber) => {
                    questions.push(cell.value);
                })
                workedRows++;
            //scorro le risposte
            }else{
                responses = [];
                row.eachCell((cell, cellNumber) => {
                    responses.push(cell.value);
                });
                document.data = {
                    questions: questions,
                    responses: responses
                }
                fileName = "IdentikitFinanziario_" + responses[4].trim().replaceAll(" ", "_") + ".pdf";
                filesName.push(fileName);
                document.path = dirname + fileName;
                pdf.create(document, options)
                    .then(result => {
                        workedRows++;
                        if(workedRows == rowsCount) {
                            zip(dirname, filesName, res);
                        }
                        console.log(result) 
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
        
});

function zip(dirname, fileNames, res) {
    // Crea un nuovo oggetto Archiver
    const archive = archiver('zip', {
        zlib: { level: 9 } // Imposta il livello di compressione massimo
    });

    // Definisci il nome del file ZIP di output
    const outputZip = fs.createWriteStream(dirname + '/output.zip');

    // Inizia l'archiviazione
    archive.pipe(outputZip);

    fileNames.forEach(name => {
        console.log(name);
        archive.file(dirname + name, { name: name });
    });
    // Aggiungi file al file ZIP

    // Puoi aggiungere anche directory intere
    // archive.directory(dirname + '/folder', 'folder');

    // Finalizza l'archivio
    archive.finalize();

    // Gestisci evento di completamento dell'archiviazione
    outputZip.on('close', function() {
        console.log('Archivio creato con successo!');
        res.download(dirname + '/output.zip');
    });

}

var server = app.listen(app.get('port'), function() {
    console.log('Express server listening on port ' + server.address().port);
});