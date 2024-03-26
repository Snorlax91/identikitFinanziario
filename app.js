var express = require('express');
var app = express()
app.set('port', 8080);
//Required package
var pdf = require("pdf-creator-node")
var fs = require('fs')
const archiver = require('archiver');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');

// Configura multer per l'upload dei file
const storage = multer.diskStorage({
    destination: function(req, file, cb) {
        cb(null, './xlsx'); // Directory di destinazione dei file
    },
    filename: function(req, file, cb) {
        // Genera un nome univoco per il file
        console.log(file.filename);
        cb(null, "IdentikitFinanziario.xlsx");
    }
});

const upload = multer({storage: storage}); // Specifica la directory di destinazione per i file uploadati


// Read HTML Template
var html = fs.readFileSync('template.html', 'utf8')

const excelFilePath = './xlsx/IdentikitFinanziario.xlsx';
const workbook = new ExcelJS.Workbook();
const dirname = "./identikit/";
var uploadedFileName = "";

var questions = [];
var responses = [];

var document = {
    html: html,
    data: {
    },
    path: ""
};

var options = { format: "A4", orientation: "portrait", border: "10mm" };

// Visualizza un modulo per caricare file
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Gestisci l'upload del file
app.post('/upload', upload.single('file'), (req, res) => {
    // req.file contiene le informazioni sul file caricato
    console.log(req.file);

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
                if (rowNumber == 1) {// La prima riga contiene le domande
                    row.eachCell((cell, cellNumber) => {
                        questions.push(cell.value);
                    })
                    workedRows++;
                    //scorro le risposte
                } else {
                    responses = [];
                    row.eachCell((cell, cellNumber) => {
                        
                        //riempio con stringhe vuote i posti mancanti che a quanto pare la libreria excel salta
                        while(cellNumber-1>responses.length)
                            responses.push("");

                        if (cell.value){
                            if(cellNumber == 1 || cellNumber == 2) {
                                // Formatta la data
                                //const formattedDate = cell.value.toLocaleDateString('it-IT', { year: 'numeric', month: '2-digit', day: '2-digit' });
                                responses.push(cell.value);
                            } else {
                                responses.push(cell.value.replace(";", "<br>"));
                            }
                        }
                        else
                            responses.push("");
                    });
                    document.data = {
                        questions: questions,
                        responses: responses
                    }
                    fileName = "IdentikitFinanziario_" + responses[6].trim().replaceAll(" ", "_") + ".pdf";
                    filesName.push(fileName);
                    document.path = dirname + fileName;
                    pdf.create(document, {
                        childProcessOptions: {
                            env: {
                                OPENSSL_CONF: '/dev/null',
                            },
                        }
                    }, options)
                        .then(result => {
                            workedRows++;
                            if (workedRows == rowsCount) {
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
                if (rowNumber == 1) {// La prima riga contiene le domande
                    row.eachCell((cell, cellNumber) => {
                        questions.push(cell.value);
                    })
                    workedRows++;
                    //scorro le risposte
                } else {
                    responses = [];
                    row.eachCell((cell, cellNumber) => {
                        if (cell.value)
                            responses.push(cell.value);
                        else
                            responses.push("");
                    });
                    document.data = {
                        questions: questions,
                        responses: responses
                    }
                    fileName = "IdentikitFinanziario_" + responses[6].trim().replaceAll(" ", "_") + ".pdf";
                    filesName.push(fileName);
                    document.path = dirname + fileName;
                    pdf.create(document, {
                        childProcessOptions: {
                            env: {
                                OPENSSL_CONF: '/dev/null',
                            },
                        }
                    }, options)
                        .then(result => {
                            workedRows++;
                            if (workedRows == rowsCount) {
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
    outputZip.on('close', function () {
        console.log('Archivio creato con successo!');
        res.download(dirname + '/output.zip');
        removeFiles(dirname);
    });

}

function removeFiles(directoryPath) {

    fs.readdir(directoryPath, (err, files) => {
        if (err) {
            console.error('Errore nella lettura della cartella:', err);
            return;
        }

        // Per ogni file nella cartella, rimuovilo
        files.forEach(file => {
            const filePath = `${directoryPath}/${file}`;
            fs.unlink(filePath, err => {
                if (err) {
                    console.error(`Errore nella rimozione del file ${filePath}:`, err);
                } else {
                    console.log(`File ${filePath} rimosso con successo`);
                }
            });
        });
    });

}

var server = app.listen(app.get('port'), function () {
    console.log('Express server listening on port ' + server.address().port);
});