const { GoogleSpreadsheet } = require('google-spreadsheet');
const credentials = require('./credentials.json');
const file = require('./file.js');

const classes = 60; // number of total classes

const getDoc = async () => {
    const doc = new GoogleSpreadsheet(file.id);

    await doc.useServiceAccountAuth({
        client_email: credentials.client_email,
        private_key: credentials.private_key.replace(/\\n/g, '\n')
    })
    await doc.loadInfo();

    const sheet = doc.sheetsByIndex[0];

    await sheet.loadCells('A1:H27'); // loads a range of cells

    // manipulate students
    try {

        for (let index = 3; index < 27; index++) {
            let situation = sheet.getCell(index, 6); // access cells using a zero-based index
            let naf = sheet.getCell(index, 7);
            let missedClasses = sheet.getCell(index, 2).value;

            // average
            let grade1 = sheet.getCell(index, 3).value;
            let grade2 = sheet.getCell(index, 4).value;
            let grade3 = sheet.getCell(index, 5).value;
            let average = (grade1 + grade2 + grade3) / 3;

            // conditions
            if (average < 50) {
                situation.value = "Reprovado por Nota";
                naf.value = 0;
            } else if (average >= 50 && average < 70) {
                situation.value = "Exame Final";
                naf.value = Math.abs((calcNaf(average)));
            } else if (average >= 70) {
                situation.value = "Aprovado";
                naf.value = 0;
            }

            const percentageMiss = ((100 / classes) * missedClasses);

            if (percentageMiss > 25) {
                situation.value = "Reprovado por Falta";
                naf.value = 0;
            }

        }
    } catch (error) {
        console.log(error);
    }
    await sheet.saveUpdatedCells();
}

const calcNaf = (avg) => {
    return Math.round(10 - avg);
}

getDoc();