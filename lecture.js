const fs = require('fs');
const ExcelJS = require('exceljs');
const readline = require('readline');
const path = require('path');

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

const phraseDebut = ">>>>>Begin Spectral Data<<<<<"; 
const phraseFin = ">>>>>End Spectral Data<<<<<"; 

rl.question('Folder Root : ', (inputFolder) => {
    const folderPath = inputFolder.trim().replace(/^["']|["']$/g, '');
    rl.question('value minimale Wavelength : ', (minValue) => {
        rl.question('value maximale Wavelength : ', (maxValue) => {
            fs.readdir(folderPath, (err, files) => {
                if (err) {
                    console.log("Erreur lors de la lecture du dossier :", err);
                    return;
                }
                const paths = files.map(file => path.join(folderPath, file));
                processFiles(paths, parseFloat(minValue), parseFloat(maxValue));
                rl.close();
            });
        });
    });
});


async function processFiles(paths, minValue, maxValue) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Données combinées');
    let currentColumn = 1; 
    for (let i = 0; i < paths.length; i++) {
        const PathFoldeur = paths[i];
        console.log(`Processing file: ${PathFoldeur}`);
        const data = fs.readFileSync(PathFoldeur, 'utf8');

        const debut = data.indexOf(phraseDebut) + phraseDebut.length;
        const fin = data.indexOf(phraseFin);

        if (debut < fin) {
            const texteExtrait = data.substring(debut, fin).trim();
            const lignes = texteExtrait.split(/\r?\n/);

            sheet.getCell(1, currentColumn).value = `Wavelength(nm) (Fichier ${i + 1})`;
            sheet.getCell(1, currentColumn + 1).value = `Intensity (Fichier ${i + 1})`;

            let row = 2;
            lignes.forEach((ligne) => {
                const valeurs = ligne.split(/\s+/);
                if (valeurs.length >= 2) {
                    const wavelength = parseFloat(valeurs[0]);
                    if (wavelength >= minValue && wavelength <= maxValue) {
                        sheet.getCell(row, currentColumn).value = valeurs[0];
                        sheet.getCell(row, currentColumn + 1).value = valeurs[1];
                        row++;
                    }
                }
            });
            currentColumn += 2;
        } else {
            console.log(`Les marqueurs ne sont pas dans l'ordre correct ou manquants dans le fichier ${PathFoldeur}.`);
        }
    }

    await workbook.xlsx.writeFile('DonneesExtraites.xlsx');
    console.log("Les données ont été écrites dans DonneesExtraites.xlsx");
}



