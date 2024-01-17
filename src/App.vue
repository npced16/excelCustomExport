<template>
  <div>
    <button @click="exportToExcel">Export to Excel</button>
  </div>
</template>

<script setup>
import ExcelJS from 'exceljs';
const backgroundColors = {
  header: '328493',
  default: 'C7C7C7',
};
const textColor = {
  special: 'FF22FF',
  white: 'FFFFFF',
  default: '000000',
};

const defaultCellStyle = {
  font: {
    name: '휴먼편지체',
    family: 2,
    size: 10,
  },
  alignment: {
    vertical: 'middle',
    horizontal: 'center',
  },
};
const headerCellStyle = {
  ...defaultCellStyle,
  font: {
    ...defaultCellStyle.font,
    bold: true,
  },
  border: {
    bottom: { style: 'medium' },
  },
};
function exportToExcel() {
  // 새로운 워크북 및 워크시트 생성
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet 1');
  const worksheet3 = workbook.addWorksheet('Sheet 3');
  worksheet.columns = [
    { header: "", key: "id", width: 10, style: defaultCellStyle },
    { header: "", key: "name", width: 32, style: defaultCellStyle },
    { header: "", key: "dob", width: 15, style: defaultCellStyle },
    { header: "", key: "number", width: 15, style: defaultCellStyle },
    { header: "", key: "sex", width: 15, style: defaultCellStyle },
  ];
  // header 
  const header = {
    id: "증례번호",
    name: "연구대상자 이니셜",
    dob: "No",
    sex: "성별",
    number: "대상자 번호",
  }
  worksheet.addRow(
    header
  )
  worksheet.getCell('D1').value = "A.기본정보"
  worksheet.mergeCells('D1:E1');


  worksheet.autoFilter = {
    from: {
      row: 2,
      column: 1
    },
    to: {
      row: 3,
      column: 13
    }
  };
  const totalData = [
    {
      id: 1,
      name: "John Doe",
      number: "2753982",
      dob: new Date(1970, 1, 1),
      sex: 1,
    }, {
      id: 2,
      name: "Jane Doe",
      number: "1232753982",
      dob: new Date(1960, 1, 1),
      sex: 2,
    }, {
      id: 3,
      name: "John Sinna",
      number: "275398ji2kop2",
      dob: new Date(1970, 12, 1),
      sex: 2,
    }, {
      id: 4,
      name: "Sinna",
      number: "6891rfgy298ry",
      dob: new Date(1970, 12, 1),
      sex: 1,
    }, {
      id: 83,
      name: "John ",
      number: "vhbwqiu289",
      dob: new Date(1970, 12, 1),
      sex: 2,
    }
  ]
  // Add an array of rows with inherited style
  // These new rows will have same styles as last row
  // and return them as array of row objects
  // const newRowsStyled = worksheet.addRows(rows, 'i');
  for (const item of totalData) {
    worksheet.addRow(item)
  }
  function setCellFill(cell, color) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: color },
    };
  }
  function setCellStyle(cell, Cellstyle) {
    if (Cellstyle != null) {
      // Cellstyle이 null이 아닌 경우에만 셀 스타일을 설정
      cell.font = { ...Cellstyle.font };
      cell.alignment = { ...Cellstyle.alignment };
      if (Cellstyle.border) {
        cell.border = { ...Cellstyle.border };
      }
    }
  }
  function setCellcolor(cell, color) {
    cell.font = { ...(cell.font), color: { argb: color } };
  }

  worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
    const headerNumber = 2;
    row.eachCell(function (cell, colNumber) {
      if (rowNumber === headerNumber) {
        setCellStyle(cell, headerCellStyle);
        row.height = 30;
        const bgColor = colNumber == 3 || colNumber == 4 ? backgroundColors.header : backgroundColors.default;
        if (bgColor == backgroundColors.header) {
          setCellcolor(cell, textColor.white)
        }
        setCellFill(cell, bgColor);
      }
      else if (rowNumber != headerNumber) {
        if (colNumber === 5 && cell.value === 1) {
          setCellcolor(cell, textColor.special);
        }
      }
    });
  });





  const temdata = [{ id: 1, name: "J4314ane Doe", dob: new Date(1965, 1, 7) }, { id: 1, name: "J4314ane Doe", dob: new Date(1965, 1, 7) }]
  dataSetting(worksheet3, temdata, 14, 2)


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
/**
 * 
 * @param {*} ws workSheet 입력 
 * @param {*} dataArr row 에 들어갈 데이터 입력
 * @param {*} chartLength  row 길이
 * @param {*} titleNumber  차트 Header 
 */
async function dataSetting(ws, dataArr, chartLength, titleNumber) {
  ws.columns = [
    { header: 'Num', key: 'id' },
    { header: 'Nom prenom', key: 'name' },
    { header: 'Date de naissance', key: 'dob' },
  ];
  for (const item of dataArr) {
    ws.addRow(item);
  }
  ws.autoFilter = {
    from: {
      row: 1,
      column: 1
    },
    to: {
      row: 5,
      column: 13
    }
  };
  ws.eachRow({ includeEmpty: true }, function (row, rowNumber) {

    row.eachCell(function (cell) {
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
}


</script>
