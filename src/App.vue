<template>
  <div>
    <button @click="exportToExcel">Export to Excel</button>
  </div>
</template>

<script setup>
import ExcelJS from 'exceljs';


function exportToExcel() {
  // 새로운 워크북 및 워크시트 생성
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet 1');
  const worksheet2 = workbook.addWorksheet('Sheet 2');
  const worksheet3 = workbook.addWorksheet('Sheet 3');
  // const sheetOne = workbook.getWorksheet('Sheet 2');
  // 열 설정
  worksheet.columns = [
    { header: "Id", key: "id", width: 10 },
    { header: "Name", key: "name", width: 32 },
    { header: "D.O.B.", key: "dob", width: 15 },
  ];

  worksheet.columns = [
    {
      header: "Id",
      key: "id",
      width: 10,
      filterButton: true,
    },
    { header: "Name", key: "name", width: 32 },
    { header: "D.O.B.", key: "dob", width: 15 },
  ];
  worksheet.addRow({
    id: 1,
    name: "John Doe",
    dob: new Date(1970, 1, 1),
    style: {
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0000FF" },
      },
    },
  });
  worksheet.addRow({ id: 2, name: "Jane Doe", dob: new Date(1965, 1, 7) });

  // worksheet.getCell("B2").font = { size: 14 };
  worksheet.addTable({
    name: "MyTable",
    ref: "A2",
    headerRow: true,
    totalsRow: true,
    style: {
      theme: "TableStyleDark1",
      // showRowStripes: true,
      showFirstColumn: true,
    },
    columns: [
      { name: "Date", totalsRowLabel: "Totals:", filterButton: true },
      { name: "Amount", totalsRowFunction: "sum", filterButton: true },
    ],
    rows: [
      [new Date("2019-07-20"), 70.1],
      [new Date("2019-07-21"), 70.6],
      [new Date("2019-07-22"), 70.1],
    ],
  });






  // 첫번째 테이블 정의
  const table1Layout = {
    name: 'table1',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: {
      theme: 'TableStyleLight1',
      showRowStripes: false,
    },
    columns: [
      {
        name: 'Date', filterButton: true, style: {
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF0000FF" },
          },
        },
      },
      { name: 'Amount', filterButton: false },
    ],
    rows: [
      [new Date('2023-01-01'), 70.10],
      [new Date('2023-01-02'), 70.60],
      [new Date('2023-01-03'), 70.10],
    ],
  }

  // 두번째 테이블 정의
  const table2Layout = {
    name: 'table2',
    ref: 'C7',
    headerRow: true,
    totalsRow: true,
    style: {
      theme: 'TableStyleLight8',
      showRowStripes: false,
    },
    columns: [
      { name: 'Date', filterButton: true },
      { name: 'Count1', filterButton: true },
      { name: 'Count2', filterButton: true },
    ],
    rows: [
      [new Date('2023-01-01'), 70.10],
      [new Date('2023-01-02'), 70.60],
      [new Date('2023-01-03'), 70.10],
    ],
  }
  worksheet3.autoFilter = {
    from: 'A1',
    to: 'A1000', // 특정 범위 설정
  };
  // Sheet One의 컬럼에 대한 스타일 설정
  worksheet2.getColumn(1).style = { numFmt: 'YYYY-MM-DD', alignment: { horizontal: 'center' } }
  worksheet2.getColumn(1).width = 20
  worksheet2.getColumn(3).width = 20

  // 두 개의 테이블을 Sheet One에 추가
  worksheet2.addTable(table1Layout)
  worksheet2.addTable(table2Layout)



  // 샘플 데이터를 추가
  worksheet.addRow({ id: 1, name: "John Doe", dob: new Date(1970, 1, 1) });

  worksheet3.columns = [
    { header: 'Num', key: 'id' },
    { header: 'Nom prenom', key: 'name' },
    { header: 'Date de naissance', key: 'dob' },
  ]
  worksheet3.addRow({ id: 1, name: "J4314ane Doe", dob: new Date(1965, 1, 7) });
  worksheet3.addRow({ id: 5, name: "Jan43143e Doe", dob: new Date(1965, 1, 7) });
  worksheet3.addRow({ id: 3, name: "Ja34211431 Doe", dob: new Date(1965, 1, 7) });
  worksheet3.addRow({ id: 5, name: "Jane 43Doe", dob: new Date(1965, 1, 7) });
  worksheet3.addRow({ id: 4, name: "Ja1431e12341 1432Doe", dob: new Date(1965, 1, 7) });
  worksheet3.addRow({ id: 3, name: "Jan134e Doe", dob: new Date(1965, 1, 7) }); worksheet3.addRow({ id: 1, name: "J4314ane Doe", dob: new Date(1965, 1, 7) });
  worksheet3.addRow({ id: 5, name: "Jan43143e Doe", dob: new Date(1965, 1, 7) });
  worksheet3.addRow({ id: 3, name: "Ja34211431 Doe", dob: new Date(1965, 1, 7) });
  worksheet3.addRow({ id: 5, name: "Jane 43Doe", dob: new Date(1965, 1, 7) });
  worksheet3.addRow({ id: 4, name: "Ja1431e12341 1432Doe", dob: new Date(1965, 1, 7) });
  worksheet3.addRow({ id: 3, name: "Jan134e Doe", dob: new Date(1965, 1, 7) });
  worksheet3.autoFilter = {
    from: {
      row: 1,
      column: 1
    },
    to: {
      row: 5,
      column: 13
    }
  };
  worksheet3.eachRow({ includeEmpty: true }, function (row, rowNumber) {
    const chartLength = 14
    const titleNumber = 2

    row.eachCell(function (cell, colNumber) {
      console.log('colNumber :>> ', cell._address, colNumber, cell);
      cell.font = {
        name: 'Arial',
        family: 2,
        bold: false,
        size: 10,
      };
      cell.alignment = {
        vertical: 'middle', horizontal: 'center'
      };

      // header 위에 부분
      if (rowNumber <= titleNumber) {
        row.height = 20;
        cell.font = {
          bold: true,
        };
      }

    });
    for (var i = 1; i < chartLength; i++) {
      if (rowNumber == 1) {
        row.getCell(i).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'C7C7C7' }
        };
      }
      row.getCell(i).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
  });
  // 파일 다운로드
  workbook.xlsx.writeBuffer()
    .then(buffer => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'output.xlsx';
      link.click();
    })
    .catch(error => {
      console.error('Error exporting to Excel:', error);
    });
}


</script>
