import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
import * as logoFile from './carlogo.js';
import { DatePipe } from '@angular/common';
@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  constructor(private datePipe: DatePipe) {

  }

  async generateExcel() {


    // const ExcelJS = await import('exceljs');
    // console.log(ExcelJS);
    // const Workbook: any = {};

  // Excel Title, Header, Data
    const title = 'Leather Sector Skill Council';
    const header = ['Year', 'Month', 'Make', 'Model', 'Quantity', 'Pct'];
    const data = [
    [2007, 1, 'Volkswagen ', 'Volkswagen Passat', 1267, 10],
    [2007, 1, 'Toyota ', 'Toyota Rav4', 819, 6.5],
    [2007, 1, 'Toyota ', 'Toyota Avensis', 787, 6.2],
    [2007, 1, 'Volkswagen ', 'Volkswagen Golf', 720, 5.7],
    [2007, 1, 'Toyota ', 'Toyota Corolla', 691, 5.4],
    [2007, 1, 'Peugeot ', 'Peugeot 307', 481, 3.8],
    [2008, 1, 'Toyota ', 'Toyota Prius', 217, 2.2],
    [2008, 1, 'Skoda ', 'Skoda Octavia', 216, 2.2],
    [2008, 1, 'Peugeot ', 'Peugeot 308', 135, 1.4],
    [2008, 2, 'Ford ', 'Ford Mondeo', 624, 5.9],
    [2008, 2, 'Volkswagen ', 'Volkswagen Passat', 551, 5.2],
    [2008, 2, 'Volkswagen ', 'Volkswagen Golf', 488, 4.6],
    [2008, 2, 'Volvo ', 'Volvo V70', 392, 3.7],
    [2008, 2, 'Toyota ', 'Toyota Auris', 342, 3.2],
    [2008, 2, 'Volkswagen ', 'Volkswagen Tiguan', 340, 3.2],
    [2008, 2, 'Toyota ', 'Toyota Avensis', 315, 3],
    [2008, 2, 'Nissan ', 'Nissan Qashqai', 272, 2.6],
    [2008, 2, 'Nissan ', 'Nissan X-Trail', 271, 2.6],
    [2008, 2, 'Mitsubishi ', 'Mitsubishi Outlander', 257, 2.4],
    [2008, 2, 'Toyota ', 'Toyota Rav4', 250, 2.4],
    [2008, 2, 'Ford ', 'Ford Focus', 235, 2.2],
    [2008, 2, 'Skoda ', 'Skoda Octavia', 225, 2.1],
    [2008, 2, 'Toyota ', 'Toyota Yaris', 222, 2.1],
    [2008, 2, 'Honda ', 'Honda CR-V', 219, 2.1],
    [2008, 2, 'Audi ', 'Audi A4', 200, 1.9],
    [2008, 2, 'BMW ', 'BMW 3-serie', 184, 1.7],
    [2008, 2, 'Toyota ', 'Toyota Prius', 165, 1.6],
    [2008, 2, 'Peugeot ', 'Peugeot 207', 144, 1.4]
  ];

    // Create workbook and worksheet
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Assessment Report');


// Add Row and formatting
//row 1
    const titleRow = worksheet.addRow([title]);
    titleRow.font = { name: 'Cambria', family: 4, size: 18, underline: 'single', bold: true,color:{argb: 'FFFF00'} };

    titleRow.alignment ={horizontal:'center',readingOrder: 'rtl',vertical:'middle'}
    titleRow.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
    titleRow.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '1F497D' }
    };
    titleRow.height= 50

    // worksheet.addRow([]);
    //row 2
    const subTitleRow = worksheet.addRow(['Assessment Report']);
    subTitleRow.font = { name: 'Cambria', family: 4, size: 16, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
    subTitleRow.alignment ={horizontal:'center',readingOrder: 'rtl',vertical:'middle'}
    subTitleRow.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
    subTitleRow.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '1F497D' }
    };
    subTitleRow.height = 40
//row3
const subTitleRow2 = worksheet.addRow(['Batch ID - 2030536']);
    subTitleRow2.font = { name: 'Cambria', family: 4, size: 15, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
    subTitleRow2.alignment ={horizontal:'center',readingOrder: 'rtl',vertical:'middle'}
    subTitleRow2.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
    subTitleRow2.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '1F497D' }
    };
    subTitleRow2.height = 40
// Add Image
    const logo = workbook.addImage({
  base64: logoFile.logoBase64,
  extension: 'png',
});

worksheet.mergeCells('J1:K3');
    worksheet.addImage(logo, 'J1:K3',);
    worksheet.mergeCells('A1:I1');
    worksheet.mergeCells('A2:I2');
    worksheet.mergeCells('A3:I3');


// row 4
  const row4 =  worksheet.addRow(['Assessment Date:']);
  row4.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
  worksheet.getColumn(1).width = 20;
  row4.font = { name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
  row4.getCell(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '1F497D' }
  };
  row4.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
  worksheet.mergeCells('B4:C4');

const colum4 = worksheet.getCell('D4');
colum4.value='Assessor Name:'
colum4.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
colum4.font ={ name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
colum4.fill= {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '1F497D' }
}
colum4.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('E4:I4');



// row 5
const row5 =  worksheet.addRow(['Request ID:']);
row5.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
worksheet.getColumn(1).width = 20;
row5.font = { name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
row5.getCell(1).fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '1F497D' }
};
row5.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('B5:C5');

const colum5 = worksheet.getCell('D5');
colum5.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
colum5.value='Assessor AR ID:'
colum5.font ={ name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
colum5.fill= {
type: 'pattern',
pattern: 'solid',
fgColor: { argb: '1F497D' }
}
colum5.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('E5:I5');


// row 6
const row6 =  worksheet.addRow(['No. of Student:']);
row6.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
worksheet.getColumn(1).width = 20;
row6.font = { name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
row6.getCell(1).fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '1F497D' }
};
row6.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('B6:C6');

const colum6 = worksheet.getCell('D6');
colum6.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
colum6.value='Start Date:'
colum6.font ={ name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
colum6.fill= {
type: 'pattern',
pattern: 'solid',
fgColor: { argb: '1F497D' }
}
colum6.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('E6:I6');

//row7
const row7 =  worksheet.addRow(['Contact Person:']);
row7.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
worksheet.getColumn(1).width = 20;
row7.font = { name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
row7.getCell(1).fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '1F497D' }
};
row7.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('B7:C7');

const colum7 = worksheet.getCell('D7');
colum7.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
colum7.value='End Date:'
colum7.font ={ name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
colum7.fill= {
type: 'pattern',
pattern: 'solid',
fgColor: { argb: '1F497D' }
}
colum7.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('E7:I7');


//row 8

const row8 =  worksheet.addRow(['Contact No:']);
row8.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
worksheet.getColumn(1).width = 20;
row8.font = { name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
row8.getCell(1).fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '1F497D' }
};
row8.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('B8:C8');

const colum8 = worksheet.getCell('D8');
colum8.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
colum8.value='Total No. of Present'
colum8.font ={ name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
colum8.fill= {
type: 'pattern',
pattern: 'solid',
fgColor: { argb: '1F497D' }
}
colum8.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('E8:I8');


//row 9

const row9 =  worksheet.addRow(['Pass:']);
row9.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
worksheet.getColumn(1).width = 20;
row9.font = { name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
row9.getCell(1).fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '1F497D' }
};
row9.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('B9:C9');

const colum9 = worksheet.getCell('D9');
colum9.alignment ={horizontal:'right',readingOrder: 'ltr',vertical:'middle'}
colum9.value='Fail'
colum9.font ={ name: 'Cambria', family: 4, size: 10, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };
colum9.fill= {
type: 'pattern',
pattern: 'solid',
fgColor: { argb: '1F497D' }
}
colum9.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
worksheet.mergeCells('E9:I9');


//row 10

const row10 = worksheet.addRow(['QP  NOS - Cutter (Footwear)  LSS/Q2301'])
row10.font = { name: 'Cambria', family: 4, size: 18, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };

row10.alignment ={horizontal:'center',readingOrder: 'rtl',vertical:'middle'}
row10.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
row10.getCell(1).fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '1F497D' }
};
row10.height= 50
worksheet.mergeCells('A10:K10');

//row 11

const row11 = worksheet.addRow(['Center Name & Address - NIFA DDU GKY, Tonk'])
row11.font = { name: 'Cambria', family: 4, size: 15, underline: 'none', bold: true,color:{argb: 'FFFFFF'} };

row11.alignment ={horizontal:'center',readingOrder: 'rtl',vertical:'middle'}
row11.getCell(1).border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}
row11.getCell(1).fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '1F497D' }
};
row11.height= 50
worksheet.mergeCells('A11:K11');
worksheet.mergeCells('J4:K9');

const  j4cell =  worksheet.getCell('J4')
j4cell.fill= {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: '1F497D' }
  }
  worksheet.mergeCells('A12:A15')
  worksheet.mergeCells('B12:B15')
// colu11

const column12 = worksheet.getCell('A12')

column12.value= 'Sl No'
column12.font = { name: 'Times New Roman', family: 4, size: 12, underline: 'none', bold: true,color:{argb: '000000'} };
column12.alignment ={horizontal:'center',readingOrder: 'rtl',vertical:'middle'}
column12.fill= {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFF00' }
  }
  column12.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}

//column13

  const column13 = worksheet.getCell('B12')
  column13.value= 'Enrollment No'
  column13.font = { name: 'Times New Roman', family: 4, size: 12, underline: 'none', bold: true,color:{argb: '000000'} };
  column13.alignment ={horizontal:'center',readingOrder: 'rtl',vertical:'middle'}
  column13.fill= {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' }
    }
    column13.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}

 worksheet.mergeCells('C12:D14')

//column14

const columnCD = worksheet.getCell('C12')
columnCD.value= '1. LSS/N2301–Carry out cutting operations'
columnCD.font = { name: 'Times New Roman', family: 4, size: 12, underline: 'none', bold: true,color:{argb: '000000'} };
columnCD.alignment ={horizontal:'distributed',readingOrder: 'rtl',vertical:'middle'}
columnCD.fill= {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFF00' }
  }
  columnCD.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}


// column 15
const column15 = worksheet.getCell('C15')
column15.value= 'Theory'
column15.font = { name: 'Times New Roman', family: 4, size: 10, underline: 'none', bold: true,color:{argb: '000000'} };
column15.alignment ={horizontal:'distributed',readingOrder: 'rtl',vertical:'middle'}
column15.fill= {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFF00' }
  }
  column15.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}




//column16

const column16 = worksheet.getCell('D15')
column16.value= 'Practical'
column16.font = { name: 'Times New Roman', family: 4, size: 10, underline: 'none', bold: true,color:{argb: '000000'} };
column16.alignment ={horizontal:'distributed',readingOrder: 'rtl',vertical:'middle'}
column16.fill= {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFF00' }
  }
  column16.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}

  worksheet.mergeCells('E12:F14')


  const columnE = worksheet.getCell('E12')
  columnE.value= 'LSS/N2302– Contribute to achieving product quality in cutting processes	'
  columnE.font = { name: 'Times New Roman', family: 4, size: 12, underline: 'none', bold: true,color:{argb: '000000'} };
  columnE.alignment ={horizontal:'distributed',readingOrder: 'rtl',vertical:'middle'}
  columnE.fill= {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF00' }
    }
    columnE.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}

/////////////
    const columnE15 = worksheet.getCell('E15')
    columnE15.value= 'Theory'
    columnE15.font = { name: 'Times New Roman', family: 4, size: 10, underline: 'none', bold: true,color:{argb: '000000'} };
    columnE15.alignment ={horizontal:'distributed',readingOrder: 'rtl',vertical:'middle'}
    columnE15.fill= {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFF00' }
  }
  columnE15.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}

  ////////

  const columnF15 = worksheet.getCell('F15')
    columnF15.value= 'Practical'
    columnF15.font = { name: 'Times New Roman', family: 4, size: 10, underline: 'none', bold: true,color:{argb: '000000'} };
    columnF15.alignment ={horizontal:'distributed',readingOrder: 'rtl',vertical:'middle'}
    columnF15.fill= {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFF00' }
  }
  columnF15.border = {  top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' }}

// // Add Header Row
//     const headerRow = worksheet.addRow(header);

// // Cell Style : Fill and Border
//     headerRow.eachCell((cell, number) => {
//   cell.fill = {
//     type: 'pattern',
//     pattern: 'solid',
//     fgColor: { argb: 'FFFFFF00' },
//     bgColor: { argb: 'FF0000FF' }
//   };
//   cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
// });
// // worksheet.addRows(data);


// // Add Data and Conditional Formatting
//     data.forEach(d => {
//   const row = worksheet.addRow(d);
//   const qty = row.getCell(5);
//   let color = 'FF99FF99';
//   if (+qty.value < 500) {
//     color = 'FF9999';
//   }

//   qty.fill = {
//     type: 'pattern',
//     pattern: 'solid',
//     fgColor: { argb: color }
//   };
// }

// );

    worksheet.getColumn(3).width = 20;
    worksheet.getColumn(4).width = 20;
    worksheet.getColumn(2).width = 30;
    worksheet.getColumn(11).width = 25;
    worksheet.getColumn(5).width = 20;
    worksheet.getColumn(6).width = 25;
    worksheet.addRow([]);


// Footer Row
    const footerRow = worksheet.addRow(['This is system generated excel sheet.']);
    footerRow.getCell(1).fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFCCFFE5' }
};
    footerRow.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

// Merge Cells
    worksheet.mergeCells(`A${footerRow.number}:F${footerRow.number}`);

// Generate Excel File with given name
    workbook.xlsx.writeBuffer().then((data: any) => {
  const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  fs.saveAs(blob, 'assessmentreport.xlsx');
});

  }
}
