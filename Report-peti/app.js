const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs');
const bodyParser = require('body-parser');

const app = express();
app.use(bodyParser.json());

app.get('/user/:userId', (req, res) => {
  const userId = req.params.userId;
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const user = data.users.find((user) => user.id === parseInt(userId));
  if (user) {
    res.send(user);
  } else {
    res.status(404).send('User not found');
  }
});

app.get('/post/:postId', (req, res) => {
  const { postId } = req.params;
  const { posts } = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const post = posts.find((post) => post.id === parseInt(postId));

  if (post) {
    res.send(post);
  } else {
    res.status(404).send('Post not found');
  }
});

app.get('/posts/:startDate/:endDate', (req, res) => {
  const { startDate, endDate } = req.params;
  const { posts } = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const filteredPosts = posts.filter((post) => {
    const postDate = new Date(post.last_update);
    return postDate >= new Date(startDate) && postDate <= new Date(endDate);
  });

  if (filteredPosts.length > 0) {
    res.send(filteredPosts);
  } else {
    res.status(404).send('No posts found between specified dates');
  }
});

app.post('/user/:userId/email', (req, res) => {
  const userId = req.params.userId;
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const user = data.users.find((user) => user.id === parseInt(userId));

  if (user) {
    user.email = req.body.email;
    fs.writeFileSync('data.json', JSON.stringify(data));
    res.send('Email updated');
  } else {
    res.status(404).send('User not found');
  }
});

app.put('/user/:userId/post', (req, res) => {
  const { userId } = req.params;
  const { users, posts } = JSON.parse(fs.readFileSync('data.json', 'utf8'));
  const { title, body } = req.body;

  const userIndex = users.findIndex((user) => user.id === parseInt(userId));

  if (userIndex !== -1) {
    const newPost = {
      id: posts.length + 1,
      userId: parseInt(userId),
      title,
      body,
      date: new Date().toISOString(),
      last_update: new Date().toISOString(),
    };
    posts.push(newPost);
    fs.writeFileSync('data.json', JSON.stringify({ users, posts }));
    res.send('Post created');
  } else {
    res.status(404).send('User not found');
  }
});

app.post('/generate-excel', (req, res) => {
  // Citaj iz jsona
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  createExcelFile(data)
    .then(() => {
      res.send('Excel file created successfully.');
    })
    .catch((error) => {
      console.error('Error:', error);
      res.status(500).send('An error occurred while generating the Excel file.');
    });
});

async function createExcelFile(data) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet 1');

  const logoImage = workbook.addImage({
    filename: 'logo.jpg', // Replace 'logo.png' with the actual filename of your logo image
    extension: 'jpg',
  });

  worksheet.addImage(logoImage, {
    tl: { col: 0.5, row: 1 }, // Adjust the coordinates (col, row) as needed to position the logo
    br: { col: 2.65, row: 3.25 },
  });

  // Predmet spajanje celija plus sadrzaj iz .json
  worksheet.mergeCells('A5:C5');
  const subject = data.profesori[0];
  const mergedCell = worksheet.getCell('A5');
  if (subject) {
    mergedCell.value = 'Predmet: ' + subject['PredmetNaziv'] + ' (' + subject['PredmetKratica'] + ')';
  } else {
    mergedCell.value = 'Predmet: N/A';
  }

  // Lorem ipsum text
  worksheet.mergeCells('A6:I11');
  worksheet.getCell('A6').value =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.';
  worksheet.getCell('A6').alignment = { wrapText: true };

  // Merge za katedre
  worksheet.mergeCells('A12:B12');
  worksheet.mergeCells('H12:I12');

  // Imenovanje za katedre
  worksheet.getCell('A12').value = 'Katedra';
  worksheet.getCell('C12').value = 'Studij';
  worksheet.getCell('D12').value = 'ak. god.';
  worksheet.getCell('E12').value = 'stud. god.';
  worksheet.getCell('F12').value = 'pocetak turnusa';
  worksheet.getCell('G12').value = 'kraj turnusa';
  worksheet.getCell('H12').value = 'broj sati predviden programom';

  //Ispis za katedru
  const katedra = data.profesori;
  const row = worksheet.getRow(13);

  let katedraSatiPred = 0;
  let katedraSatiVjez = 0;
  let katedraSatiSem = 0;
  data.profesori.forEach((professor) => {
    katedraSatiPred += professor.PlaniraniSatiPredavanja;
    katedraSatiSem += professor.PlaniraniSatiSeminari;
    katedraSatiVjez += professor.PlaniraniSatiVjezbe;
  })

  worksheet.mergeCells(`A13:B13`);
  worksheet.getCell('A13').value = katedra[0]['Katedra'];
  worksheet.getCell('C13').value = katedra[0]['Studij'];
  worksheet.getCell('D13').value = katedra[0]['SkolskaGodinaNaziv'];
  worksheet.getCell('E13').value = katedra[0]['PKSkolskaGodina'];
  worksheet.getCell('F13').value = katedra[0]['PocetakTurnusa'];
  worksheet.getCell('G13').value = katedra[0]['KrajTurnusa'];
  worksheet.mergeCells('H13:I13');
  worksheet.getCell('H13').value =
    'P: ' + 
    katedraSatiPred + ' ' +
    'S: ' + 
    katedraSatiSem + ' ' +
    'V: ' +
    + katedraSatiVjez;

  row.alignment = { horizontal: 'left' };
  worksheet.getCell('H13').alignment = { horizontal: 'center' };

  // Vanjski borderi
  worksheet.getCell('A13').border = { bottom: { style: 'medium' }, right: {style: 'thin'} };
  worksheet.getCell('B13').border = { bottom: { style: 'medium' }, right: {style: 'thin'} };
  worksheet.getCell('C13').border = { bottom: { style: 'medium' }, right: {style: 'thin'} };
  worksheet.getCell('D13').border = { bottom: { style: 'medium' }, right: {style: 'thin'} };
  worksheet.getCell('E13').border = { bottom: { style: 'medium' }, right: {style: 'thin'} };
  worksheet.getCell('F13').border = { bottom: { style: 'medium' }, right: {style: 'thin'} };
  worksheet.getCell('G13').border = { bottom: { style: 'medium' }, right: {style: 'thin'} };
  worksheet.getCell('H13').border = { bottom: { style: 'medium' }, right: {style: 'thin'} };
  worksheet.getCell('I13').border = { bottom: { style: 'medium' }, right: {style: 'medium'} };


  // Spajanje celije za parofesore
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

  // Imenovanje celije za profesore
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

  // Formatiranje headera
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

  // Visina headera
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

  // Racunanje sati
  let sumSatiNastavePred = 0;
  let sumSatiNastaveSem = 0;
  let sumSatiNastaveVjezbe = 0;
  let sumBrutoIznosPred = 0;
  let sumBrutoIznosSem = 0;
  let sumBrutoIznosVjezbe = 0;

  // Bruto racunanje
  let sumBrutoPred = 0;
  let sumBrutoSem = 0;
  let sumBrutoVjezbe = 0;

  // Sati racunanje
  let sumSatiPred = 0;
  let sumSatiSem = 0;
  let sumSatiVjezbe = 0;


  let totalSum;
  // Sami sadrzaj/podaci iz jsona
  data.profesori.forEach((professor, index) => {
    const rowNumber = 16 + index + 1;
    worksheet.getCell(`A${rowNumber}`).value = index + 1;
    worksheet.getCell(`B${rowNumber}`).value = professor['NastavnikSuradnikNaziv'];
    worksheet.getCell(`C${rowNumber}`).value = professor['Zvanje'];
    worksheet.getCell(`D${rowNumber}`).value = professor['NazivNastavnikStatus'];
    worksheet.getCell(`E${rowNumber}`).value = professor['PlaniraniSatiPredavanja'];
    worksheet.getCell(`F${rowNumber}`).value = professor['PlaniraniSatiSeminari'];
    worksheet.getCell(`G${rowNumber}`).value = professor['PlaniraniSatiVjezbe'];
    worksheet.getCell(`H${rowNumber}`).value = professor['NormaPlaniraniSatiPredavanja'];
    worksheet.getCell(`I${rowNumber}`).value = professor['NormaPlaniraniSatiSeminari'];
    worksheet.getCell(`J${rowNumber}`).value = professor['NormaPlaniraniSatiVjezbe'];
    worksheet.getCell(`K${rowNumber}`).value = professor['NormaPlaniraniSatiPredavanja'] * professor['PlaniraniSatiPredavanja'];
    worksheet.getCell(`L${rowNumber}`).value = professor['NormaPlaniraniSatiSeminari'] * professor['PlaniraniSatiSeminari'];
    worksheet.getCell(`M${rowNumber}`).value = professor['NormaPlaniraniSatiVjezbe'] * professor['PlaniraniSatiVjezbe'];
    worksheet.getCell(`N${rowNumber}`).value = worksheet.getCell(`K${rowNumber}`).value + worksheet.getCell(`L${rowNumber}`).value + worksheet.getCell(`M${rowNumber}`).value;

    const sum = worksheet.getCell(`K${rowNumber}`).value + worksheet.getCell(`L${rowNumber}`).value + worksheet.getCell(`M${rowNumber}`).value;
    worksheet.getCell(`N${rowNumber}`).value = sum ;
    worksheet.getCell(`N${rowNumber}`).border = { right: {style: 'medium'}} ;

    sumSatiNastavePred += professor['PlaniraniSatiPredavanja'];
    sumSatiNastaveSem += professor['PlaniraniSatiSeminari']
    sumSatiNastaveVjezbe += professor['PlaniraniSatiVjezbe'];

    sumBrutoIznosPred += professor['NormaPlaniraniSatiPredavanja'];
    sumBrutoIznosSem += professor['NormaPlaniraniSatiSeminari'];
    sumBrutoIznosVjezbe += professor['NormaPlaniraniSatiVjezbe'];

    sumBrutoPred = professor['NormaPlaniraniSatiPredavanja'] * professor['PlaniraniSatiPredavanja'];
    sumBrutoSem = professor['NormaPlaniraniSatiSeminari'] * professor['PlaniraniSatiSeminari'];
    sumBrutoVjezbe = professor['NormaPlaniraniSatiVjezbe'] * professor['PlaniraniSatiVjezbe'];

    totalSum = sum;

    // Formatiranje za celije
    worksheet.getRow(rowNumber).eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
  });

  // Sirine stupaca A do J, nakon J su standardni

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

  const totalRowNumber = 16 + data.profesori.length + 1;

  worksheet.mergeCells(`A${totalRowNumber}:C${totalRowNumber}`);
  worksheet.getCell(`A${totalRowNumber}`).value = 'Ukupno';
  worksheet.getCell(`A${totalRowNumber}`).alignment = { horizontal: 'center' };

  // Ne radi for each moram hardkodira?????
  worksheet.getCell(`A${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`D${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`E${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`F${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`G${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`H${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`I${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`J${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`K${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`L${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`M${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };
  worksheet.getCell(`N${totalRowNumber}`).border = { bottom: { style: 'medium' }, right: {style: 'medium'}, top: {style: 'medium'}, right: {style: 'medium'} };

  // Racunanje sati
  worksheet.getCell(`E${totalRowNumber}`).value = {
    formula: `SUM(E17:E${totalRowNumber - 1})`,
    result: sumSatiPred,
  };
  worksheet.getCell(`F${totalRowNumber}`).value = {
    formula: `SUM(F17:F${totalRowNumber - 1})`,
    result: sumSatiSem,
  };
  worksheet.getCell(`G${totalRowNumber}`).value = {
    formula: `SUM(G17:G${totalRowNumber - 1})`,
    result: sumSatiVjezbe,
  };

  //Bruto satnica
  worksheet.getCell(`H${totalRowNumber}`).value = {
    formula: `SUM(H17:H${totalRowNumber - 1})`,
    result: sumBrutoIznosPred,
  };
  worksheet.getCell(`I${totalRowNumber}`).value = {
    formula: `SUM(I17:I${totalRowNumber - 1})`,
    result: sumBrutoIznosSem,
  };
  worksheet.getCell(`J${totalRowNumber}`).value = {
    formula: `SUM(J17:J${totalRowNumber - 1})`,
    result: sumBrutoIznosVjezbe,
  };

  // Bruto iznosi
  worksheet.getCell(`K${totalRowNumber}`).value = {
    formula: `SUM(K17:K${totalRowNumber - 1})`,
    result: sumBrutoPred,
  };
  worksheet.getCell(`L${totalRowNumber}`).value = {
    formula: `SUM(L17:L${totalRowNumber - 1})`,
    result: sumBrutoSem,
  };
  worksheet.getCell(`M${totalRowNumber}`).value = {
    formula: `SUM(M17:M${totalRowNumber - 1})`,
    result: sumBrutoVjezbe,
  };

  //Ukupan iznos
  worksheet.getCell(`N${totalRowNumber}`).value = {
    formula: `SUM(K${totalRowNumber}:M${totalRowNumber})`,
    result: totalSum,
  };


  let dekani = data.dekani;

  worksheet.mergeCells('A28:C29');
  worksheet.mergeCells('A34:C35');
  worksheet.mergeCells('J34:L35');
  worksheet.getCell('A28').value = {
    richText: [
      { text: 'Prodekanica za nastavu i studentska pitanja\n' },
      { text: `Prof. dr. sc. ${dekani[0].ImePrezime}` },
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
      { text: `Prof. dr. sc. ${dekani[1].ImePrezime}` },
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
      { text: `Prof. dr. sc. ${dekani[2].ImePrezime}` },
    ],
  };
  worksheet.getCell('J34').alignment = {
    vertical: 'middle',
    horizontal: 'left',
    wrapText: true,
  };

    // Spremanje u datoteku
    await workbook.xlsx.writeFile('output.xlsx');
    console.log('Excel file created successfully.');
  }

// Slušanje na odabranom portu
app.listen(3001, () => {
  console.log('API server je pokrenut na portu 3001.');
});

