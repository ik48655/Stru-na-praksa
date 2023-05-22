const express = require('express');
const fs = require('fs');
const ExcelJS = require('exceljs');
const bodyParser = require('body-parser');

const app = express();
app.use(bodyParser.json());

async function createExcelFile() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');

  const logoImage = fs.readFileSync('logo.png');
  const logoBase64 = logoImage.toString('base64');

  // Read data from data.json
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const imageId = workbook.addImage({
    buffer: imageFile,
    extension: 'png',
  });

  const logo = workbook.addImage({
    base64: logoBase64,
    extension: 'png',
  });

  worksheet.addImage(logo, {
    tl: { col: 0.5, row: 0.5 },  // Specify the top left cell for the image
    ext: { width: 100, height: 100 },  // Set the dimensions of the image
  });

  // Predmet spajanje celija plus sadrzaj iz .json
  worksheet.mergeCells('A5:C5');
  const subject = data.subjects[0];
  const mergedCell = worksheet.getCell('A5');
  if (subject) {
    mergedCell.value = 'Predmet: ' + subject.name + ' (' + subject.code + ')';
  } else {
    mergedCell.value = 'Predmet: N/A';
  }

  // Lorem ipsum text
  worksheet.mergeCells('A6:I11');
  worksheet.getCell('A6').value = {
    richText: [
      { text: 'NALOG ZA ISPLATU\n', font: { bold: true, size: 18 } },
      {
        text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.',
      },
    ],
  };

  worksheet.getCell('A6').alignment = {
    vertical: 'middle',
    horizontal: 'center',
    wrapText: true,
  };

  // Merge cells for katedre
  worksheet.mergeCells('A12:B12');
  worksheet.mergeCells('H12:I12');

  // Naming cells for katedre
  worksheet.getCell('A12').value = 'Katedra';
  worksheet.getCell('C12').value = 'Studij';
  worksheet.getCell('D12').value = 'ak. god.';
  worksheet.getCell('E12').value = 'stud. god.';
  worksheet.getCell('F12').value = 'pocetak turnusa';
  worksheet.getCell('G12').value = 'kraj turnusa';
  worksheet.getCell('H12').value = 'broj sati predviden programom';

  data.katedre.forEach((katedra, index) => {
    const rowNumber = 13 + index;
    const row = worksheet.getRow(rowNumber);

    worksheet.mergeCells(`A${rowNumber}:B${rowNumber}`);
    worksheet.getCell(`A${rowNumber}`).value = katedra['ime'];
    worksheet.getCell(`C${rowNumber}`).value = katedra['studij'];
    worksheet.getCell(`D${rowNumber}`).value = katedra['ak. god.'];
    worksheet.getCell(`E${rowNumber}`).value = katedra['stud. god.'];
    worksheet.getCell(`F${rowNumber}`).value = katedra['pocetak turnusa'];
    worksheet.getCell(`G${rowNumber}`).value = katedra['kraj turnusa'];
    worksheet.mergeCells(`H${rowNumber}:I${rowNumber}`);
    worksheet.getCell(`H${rowNumber}`).value =
      ' P:' +
      katedra['pred'] +
      ' S:' +
      +katedra['sem'] +
      ' V:' +
      +katedra['vjez'];

    // Align text to the left
    row.alignment = { horizontal: 'left' };
    worksheet.getCell(`H${rowNumber}`).alignment = { horizontal: 'center' };
    worksheet.getCell(`A${rowNumber}`).alignment = { horizontal: 'center' };

    worksheet.getRow(rowNumber).eachCell((cell) => {
      cell.border = {
        top: { style: 'medium' },
        left: { style: 'thin' },
        bottom: { style: 'medium' },
        right: { style: 'thin' },
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
  });

  // Merge cells for professors
  worksheet.mergeCells('A15:A16');
  worksheet.mergeCells('B15:B16');
  worksheet.mergeCells('C15:C16');
  worksheet.mergeCells('D15:D16');
  worksheet.mergeCells('E15:G15');
  worksheet.mergeCells('H15:H16');
  worksheet.mergeCells('I15:I16');
  worksheet.mergeCells('J15:J16');
  worksheet.mergeCells('K15:M15');
  worksheet.mergeCells('N15:N16');

  // Naming cells for professors
  worksheet.getCell('A15').value = 'Redni broj';
  worksheet.getCell('B15').value = 'Ime i Prezime';
  worksheet.getCell('C15').value = 'Zvanje';
  worksheet.getCell('D15').value = 'Status';
  worksheet.getCell('E15').value = 'Sati Nastave';
  worksheet.getCell('E16').value = 'pred';
  worksheet.getCell('F16').value = 'sem';
  worksheet.getCell('G16').value = 'vjez';
  worksheet.getCell('H15').value = 'Bruto satnica predavanja (EUR)';
  worksheet.getCell('I15').value = 'Bruto satnica seminari (EUR)';
  worksheet.getCell('J15').value = 'Bruto satnica vjezbe (EUR)';
  worksheet.getCell('K15').value = 'Bruto iznos';
  worksheet.getCell('K16').value = 'pred';
  worksheet.getCell('L16').value = 'sem';
  worksheet.getCell('M16').value = 'vjez';
  worksheet.getCell('N15').value = 'Ukupno za isplatu (EUR)';

  // Format column headers
  const headerRows = [15, 16, 12];

  headerRows.forEach((rowNumber) => {
    const row = worksheet.getRow(rowNumber);
    row.font = { bold: true };
    row.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    };
    row.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'E7E7E7' },
      };
      cell.border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        bottom: { style: 'medium' },
        right: { style: 'medium' },
      };
    });
  });

  // Increase height of the header row
  const cellHeights = {
    A12: 50,
    D16: 100,
    E16: 100,
    F16: 100,
  };

  Object.entries(cellHeights).forEach(([cellRef, height]) => {
    const cell = worksheet.getCell(cellRef);
    const row = worksheet.getRow(cell.row);
    row.height = height;
  });
  // ...

  // Add data rows
  let rowNumber = 17;
  data.profesori.forEach((professor, index) => {
    worksheet.getCell(`A${rowNumber}`).value = index + 1;
    worksheet.getCell(`B${rowNumber}`).value =
      professor['NastavnikSuradnikNaziv'];
    worksheet.getCell(`C${rowNumber}`).value = professor['Titula'];
    worksheet.getCell(`D${rowNumber}`).value =
      professor['NazivNastavnikStatus'];
    worksheet.getCell(`E${rowNumber}`).value =
      professor['PlaniraniSatiPredavanja'];
    worksheet.getCell(`F${rowNumber}`).value =
      professor['PlaniraniSatiSeminari'];
    worksheet.getCell(`G${rowNumber}`).value = professor['PlaniraniSatiVjezbe'];
    worksheet.getCell(`H${rowNumber}`).value =
      professor['NormaPlaniraniSatiPredavanja'];
    worksheet.getCell(`I${rowNumber}`).value =
      professor['NormaPlaniraniSatiSeminari'];
    worksheet.getCell(`J${rowNumber}`).value =
      professor['NormaPlaniraniSatiVjezbe'];
    worksheet.getCell(`K${rowNumber}`).value =
      professor['NormaPlaniraniSatiPredavanja'] *
      professor['PlaniraniSatiPredavanja'];
    worksheet.getCell(`L${rowNumber}`).value =
      professor['NormaPlaniraniSatiSeminari'] *
      professor['PlaniraniSatiSeminari'];
    worksheet.getCell(`M${rowNumber}`).value =
      professor['NormaPlaniraniSatiVjezbe'] * professor['PlaniraniSatiVjezbe'];
    worksheet.getCell(`N${rowNumber}`).value =
      worksheet.getCell(`K${rowNumber}`).value +
      worksheet.getCell(`L${rowNumber}`).value +
      worksheet.getCell(`M${rowNumber}`).value;

    // Apply specific formatting to cells
    worksheet.getRow(rowNumber).eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    rowNumber++;
  });

  // Bruto racunanje
  let sumBrutoPred = 0;
  let sumBrutoSem = 0;
  let sumBrutoVjezbe = 0;

  // Sati racunanje
  let sumSatiPred = 0;
  let sumSatiSem = 0;
  let sumSatiVjezbe = 0;

  let totalSum = 0;

  worksheet.mergeCells(`A${rowNumber}:C${rowNumber}`);
  worksheet.getCell(`A${rowNumber}`).value = 'UKUPNO';
  // Racunanje sati
  worksheet.getCell(`E${rowNumber}`).value = {
    formula: `SUM(E17:E${rowNumber - 1})`,
    result: sumSatiPred,
  };
  worksheet.getCell(`F${rowNumber}`).value = {
    formula: `SUM(F17:F${rowNumber - 1})`,
    result: sumSatiSem,
  };
  worksheet.getCell(`G${rowNumber}`).value = {
    formula: `SUM(G17:G${rowNumber - 1})`,
    result: sumSatiVjezbe,
  };

  // Bruto iznosi
  worksheet.getCell(`K${rowNumber}`).value = {
    formula: `SUM(K17:K${rowNumber - 1})`,
    result: sumBrutoPred,
  };
  worksheet.getCell(`L${rowNumber}`).value = {
    formula: `SUM(L17:L${rowNumber - 1})`,
    result: sumBrutoSem,
  };
  worksheet.getCell(`M${rowNumber}`).value = {
    formula: `SUM(M17:M${rowNumber - 1})`,
    result: sumBrutoVjezbe,
  };

  //Ukupan iznos
  worksheet.getCell(`N${rowNumber}`).value = {
    formula: `SUM(K${rowNumber}:M${rowNumber})`,
    result: totalSum,
  };

  worksheet.getRow(rowNumber).eachCell((cell) => {
    cell.border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' },
      right: { style: 'medium' },
    };
    cell.alignment = { horizontal: 'center' };
    cell.font = { bold: true };
  });

  worksheet.getCell(`D${rowNumber}`).border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' },
  };
  worksheet.getCell(`H${rowNumber}`).border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' },
  };
  worksheet.getCell(`I${rowNumber}`).border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' },
  };
  worksheet.getCell(`J${rowNumber}`).border = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' },
  };

  //setting width
  worksheet.columns.forEach((column, index) => {
    if (index === 0) {
      column.width = 6;
    } else if (index === 1) {
      column.width = 18.43;
    } else if (index === 2 || index === 3) {
      column.width = 21.14;
    } else if (index === 4) {
      column.width = 6.14;
    } else if (index === 5) {
      column.width = 7.86;
    } else if (index === 6) {
      column.width = 8.14;
    } else if (index === 7) {
      column.width = 10.14;
    } else if (index === 8) {
      column.width = 10;
    } else if (index === 9) {
      column.width = 10.14;
    } else {
      column.width = 8.43;
    }
  });

  //Footer
  worksheet.mergeCells('A28:C29');
  worksheet.mergeCells('A34:C35');
  worksheet.mergeCells('J34:L35');
  worksheet.getCell('A28').value = {
    richText: [
      { text: 'Prodekanica za nastavu i studentska pitanja\n' },
      { text: `Prof. dr. sc. ${data.dekani[0].ImePrezime}` },
    ],
  };
  worksheet.getCell('A28').alignment = {
    vertical: 'middle',
    horizontal: 'left',
    wrapText: true,
  };

  worksheet.getCell('A34').value = {
    richText: [
      { text: 'Prodekan za financije i upravljanje\n' },
      { text: `Prof. dr. sc. ${data.dekani[1].ImePrezime}` },
    ],
  };
  worksheet.getCell('A34').alignment = {
    vertical: 'middle',
    horizontal: 'left',
    wrapText: true,
  };

  worksheet.getCell('J34').value = {
    richText: [
      { text: 'Dekan\n' },
      { text: `Prof. dr. sc. ${data.dekani[2].ImePrezime}` },
    ],
  };
  worksheet.getCell('J34').alignment = {
    vertical: 'middle',
    horizontal: 'left',
    wrapText: true,
  };

  // Save the workbook
  await workbook.xlsx.writeFile('output.xlsx');
  console.log('Excel file created successfully.');
}

app.get('/professor/:professorId', (req, res) => {
  const professorId = req.params.professorId;
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const professor = data.professors.find(
    (professor) => professor.id === parseInt(professorId)
  );
  if (professor) {
    res.send(professor);
  } else {
    res.status(404).send('Professor not found');
  }
});

app.get('/katedra/:katedraId', (req, res) => {
  const katedraId = req.params.katedraId;
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const katedra = data.katedre.find(
    (katedra) => katedra.id === parseInt(katedraId)
  );
  if (katedra) {
    res.send(katedra);
  } else {
    res.status(404).send('Katedra not found');
  }
});

app.post('/create-excel', (req, res) => {
  // Citaj iz jsona
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  createExcelFile(data)
    .then(() => {
      res.send('Excel file created successfully.');
    })
    .catch((error) => {
      console.error('Error:', error);
      res
        .status(500)
        .send('An error occurred while generating the Excel file.');
    });
});

const port = 3000;

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
