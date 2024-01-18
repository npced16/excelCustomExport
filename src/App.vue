<template>
  <div>
    <button @click="exportToExcel">Export to Excel</button>
  </div>
</template>

<script setup>
import ExcelJS from 'exceljs';
const backgroundColors = {
  start: 'FFCCFF',
  header: '328493',
  default: 'C7C7C7',
};
const textColor = {
  red: 'FF0000',
  white: 'FFFFFF',
  default: '000000',
};

const defaultCellStyle = {
  font: {
    name: '맑은 고딕',
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
  // worksheet.columns = [
  //   { header: "", key: "id", width: 10, style: defaultCellStyle },
  //   { header: "", key: "name", width: 32, style: defaultCellStyle },
  //   { header: "", key: "dob", width: 15, style: defaultCellStyle },
  //   { header: "", key: "number", width: 15, style: defaultCellStyle },
  //   { header: "", key: "sex", width: 15, style: defaultCellStyle },
  //   { header: "", key: "sumJ2_4", width: 15, style: defaultCellStyle },
  //   { header: "", key: "j1_Z", width: 15, style: defaultCellStyle },
  //   { header: "", key: "j2_Z", width: 15, style: defaultCellStyle },
  //   { header: "", key: "j3_Z", width: 15, style: defaultCellStyle },
  //   { header: "", key: "j4_Z", width: 15, style: defaultCellStyle },
  //   { header: "", key: "시행1", width: 15, style: defaultCellStyle },
  //   { header: "", key: "시행2", width: 15, style: defaultCellStyle },
  //   { header: "", key: "시행3", width: 15, style: defaultCellStyle },
  //   { header: "", key: "j5_Z", width: 15, style: defaultCellStyle },
  //   { header: "", key: "j6_Z", width: 15, style: defaultCellStyle },
  //   { header: "", key: "remind", width: 15, style: defaultCellStyle },
  //   { header: "", key: "j7_Z", width: 15, style: defaultCellStyle },
  //   { header: "", key: "agree", width: 15, style: defaultCellStyle },
  //   { header: "", key: "none", width: 15, style: defaultCellStyle },
  //   { header: "", key: "j8_Z", width: 15, style: defaultCellStyle },
  // ];
  const cognitiveHearingHeaders = {
    A1: "증례번호(IRB 증례기록지용)",
    B1: "연구대상자 이니셜(IRB 증례기록지용)",
    C1: "NO",
    D1: "대상자 번호",
    E1: "성별",
    F1: "검사일자",
    G1: "나이",
    H1: "교육년수",
    I1: "유입경로",
    J1: "기타",
    K1: "추후 참여 여부*다른 연구 및 2~5년도 연구 계속 참여 여부",
    L1: "종단 연구 의견",
    M1: "난청유무양이(PTA평균)",
    N1: "J1. 언어유창성 검사 : 동물 범주",
    O1: "J2. 보스톤 이름대기 검사",
    P1: "J3. MMSE-KC",
    Q1: "J3. MMSE-KC(정신과에서 측정)",
    R1: "J4. 단어목록 기억검사",
    S1: "J4. 시행(1)",
    T1: "J4. 시행 (2)",
    U1: "J4. 시행(3)",
    V1: "J5. 구성행동 검사",
    W1: "J6. 단어목록 회상검사",
    X1: "J6.회상률(%)",
    Y1: "J7. 단어목록 재인검사",
    Z1: "예",
    AA1: "아니오",
    AB1: "J8. 구성회상 검사",
    AC1: "J9.길찾기 A",
    AD1: "길찾기 B",
    AE1: "J가.단어페이지",
    AF1: "J가.색깔페이지",
    AG1: "J가.색깔 - 단어 페이지",
    AH1: "총점 I(J1+J2+J4+J5+J6+J7)",
    AI1: "총점 II(J1+J2+J4+J5+J6+J7+J8)",
    AJ1: "J1.z-score",
    AK1: "J2.z-score",
    AL1: "J3.z-score",
    AM1: "J4.z-score",
    AN1: "J4(1).z-score",
    AO1: "J4(2).z-score",
    AP1: "J4(3).z-score",
    AQ1: "J5.z-score",
    AR1: "J6.z-score",
    AS1: "J6(회상률).z-score",
    AT1: "J7(통합).z-score",
    AU1: "J7(예).z-score",
    AV1: "J7(아니오).z-score",
    AW1: "J8.z-score",
    AX1: "J9(A).z-score",
    AY1: "J9(B).z-score",
    AZ1: "J가(단어).z-score",
    BA1: "J가(색깔).z-score",
    BB1: "J가(색깔-단어).z-score",
    BC1: "총점 I.z-score",
    BD1: "총점 II.z-score",
    BE1: "A. 전화 사용능력",
    BF1: "B. 물건사기",
    BG1: "C. 음식준비하기",
    BH1: "D. 집안일 하기",
    BI1: "E. 빨래하기",
    BJ1: "F. 교통수단 이용",
    BK1: "G. 약 복용하기",
    BL1: "H. 돈 관리 능력",
    BM1: "GDS-KR(노인우울 척도)",
    BN1: "C2. BDS-ADL총점 (정신과에서 검사)",
    BO1: "C4. SBT-K 총점",
    BP1: "C4. z-score",
    BQ1: "GDS(Global Deterioration scale)",
    BR1: "(우)골도 250hz",
    BS1: "(우)골도 500hz",
    BT1: "(우)골도 1khz",
    BU1: "(우)골도 2khz",
    BV1: "(우)골도 4khz",
    BW1: "(우)골도 8khz",
    BX1: "(우)기도250hz",
    BY1: "(우)기도500hz",
    BZ1: "(우)기도1khz",
    CA1: "(우)기도2khz",
    CB1: "(우)기도4khz",
    CC1: "(우)기도8khz",
    CD1: "(좌)골도 250hz",
    CE1: "(좌)골도 500hz",
    CF1: "(좌)골도 1khz",
    CG1: "(좌)골도 2khz",
    CH1: "(좌)골도 4khz",
    CI1: "(좌)골도 8khz",
    CJ1: "(좌)기도 250hz",
    CK1: "(좌)기도 500hz",
    CL1: "(좌)기도 1khz",
    CM1: "(좌)기도 2khz",
    CN1: "(좌)기도 4khz",
    CO1: "(좌)기도 8khz",
    CP1: "Rt.-SRT(speech recognition threshold)",
    CQ1: "Lt.-SRT(speech recognition threshold)",
    CR1: "Rt. Discrimination",
    CS1: "Rt. dbHL(m)",
    CT1: "Lt. discrimination",
    CU1: "Lt. dbHL(m)",
    CV1: "Rt.-SRT(speech recognition threshold)_Aided",
    CW1: "Lt.-SRT(speech recognition threshold)_Aided",
    CX1: "Rt. Discrimination_Aided",
    CY1: "Rt. dbHL(m)_Aided",
    CZ1: "Lt. discrimination_Aided",
    DA1: "Lt. dbHL(m)_Aided",
    DB1: "임피던스_Rt. Side",
    DC1: "임피던스_Lt. side",
    DD1: "보청기착용 우측250khz",
    DE1: "보청기착용 우측500khz",
    DF1: "보청기착용 우측1khz",
    DG1: "보청기착용 우측2khz",
    DH1: "보청기착용 우측4khz",
    DI1: "보청기착용 우측8khz",
    DJ1: "보청기 좌측250khz",
    DK1: "보청기 좌측500khz",
    DL1: "보청기 좌측1khz",
    DM1: "보청기 좌측2khz",
    DN1: "보청기 좌측4khz",
    DO1: "보청기 좌측8khz",
    DP1: "양이(PTA평균)",
    DQ1: "우측(PTA평균)",
    DR1: "좌측(PTA평균)"
  };
  const columns = Object.keys(cognitiveHearingHeaders).map((key, index) => ({
    header: '',
    key: key,
    width: 15,
    style: defaultCellStyle
  }));
  console.log('columns :>> ', columns);
  worksheet.columns = columns
  // header 

  worksheet.addRow(
    cognitiveHearingHeaders
  )
  worksheet.getCell('D1').value = "A.기본정보"
  worksheet.mergeCells('D1:E1');
  function getRandomValue(key) {
    if (key === "id") {
      // "id" 또는 "number" 키에 대해서는 랜덤 알파벳 3개 생성
      const randomAlphabet = () => String.fromCharCode(Math.floor(Math.random() * 26) + 65);
      return randomAlphabet() + randomAlphabet() + randomAlphabet();
    } else if (key === "sumJ2_4") {
      return (Math.random() * 2) - 1;
    } else if (key === "sex") {
      return Math.floor(Math.random() * 2) + 1;
    } else if (key === " 시행1") {
      return (Math.random() * 2) - 1 > 0.5 ? (Math.random() * 2) - 1 : '';
    }
    else {
      // 다른 키에 대해서는 기본적으로 키 + '_value'를 리턴
      return (Math.random() * 2) - 1 > 0.5 ? (Math.random() * 2) - 1 : "NA";
    }
  }
  const randomObjects = Array.from({ length: 100 }, () =>
    Object.fromEntries(
      Object.entries(cognitiveHearingHeaders).map(([key]) => [key, getRandomValue(key)])
    )
  );
  console.log(' :>> ', randomObjects.length + 2);
  // const totalData = [
  //   {
  //     id: 1,
  //     name: "John Doe",
  //     number: "2753982",
  //     dob: new Date(1970, 1, 1),
  //     j1_Z: 20,
  //     j2_Z: 31,
  //     j3_Z: 31,
  //     sex: 1,
  //   }, {
  //     id: 2,
  //     name: "Jane Doe",
  //     number: "1232753982",
  //     dob: new Date(1960, 1, 1),
  //     j1_Z: 10,

  //     sex: 2,
  //   }, {
  //     id: 3,
  //     name: "John Sinna",
  //     number: "275398ji2kop2",
  //     dob: new Date(1970, 12, 1),
  //     j1_Z: 20,
  //     j3_Z: 31,

  //     sex: 2,
  //   }, {
  //     id: 4,
  //     name: "Sinna",
  //     number: "6891rfgy298ry",
  //     dob: new Date(1970, 12, 1),
  //     j1_Z: 40,

  //     sex: 1,
  //   }, {
  //     id: 83,
  //     name: "John ",
  //     number: "vhbwqiu289",
  //     dob: new Date(1970, 12, 1),
  //     j2_Z: 31,
  //     j3_Z: 31,
  //     sex: 2,
  //   }
  // ]
  // Add an array of rows with inherited style
  // These new rows will have same styles as last row
  // and return them as array of row objects
  // const newRowsStyled = worksheet.addRows(rows, 'i');
  for (const item of randomObjects) {
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
  function setCellColor(cell, color) {
    cell.font = { ...(cell.font), color: { argb: color } };
  }
  const lastRownumber = randomObjects.length + 2
  const headerNumber = 2;
  worksheet.fillFormula(`F3:F${lastRownumber}`, 'SUM(G3:J3)');
  worksheet.addConditionalFormatting({
    ref: `F3:F${lastRownumber}`,
    rules: [
      {
        type: 'cellIs',
        formulae: ['=0'],
        operator: "lessThan",
        // formulae: ['SUM(G3:J3)<0'],
        style: { font: { color: { argb: textColor.red } } },
      }
    ]
  })


  worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
    row.eachCell(function (cell, colNumber) {
      if (rowNumber === headerNumber) {
        setCellStyle(cell, headerCellStyle);
        row.height = 41;
        let bgColor = backgroundColors.default
        switch (colNumber) {
          case 1:
          case 2:
            bgColor = backgroundColors.start;
            break;
          case 3:
          case 4:
          case 5:
            bgColor = backgroundColors.header;
            setCellColor(cell, textColor.white);
            break;
          case 6:
            setCellColor(cell, textColor.red);
            break;
          default:
        }
        setCellFill(cell, bgColor);
      }
      else if (rowNumber != headerNumber) {
        if (colNumber === 6 && cell.value < 0) {
          setCellColor(cell, textColor.red);
        }
      }
    });
    // lastRownumber = rowNumber

  })



  worksheet.autoFilter = {
    from: {
      row: headerNumber,
      column: 1
    },
    to: {
      row: headerNumber,
      column: worksheet.columns.length
    }
  };

  const temdata = [{ id: 1, name: "J4314ane Doe", dob: new Date(1965, 1, 7) }, { id: 1, name: "J4314ane Doe", dob: new Date(1965, 1, 7) }]
  dataSetting(worksheet3, temdata, 14, 2)


  // 파일 다운로드
  workbook.xlsx.writeBuffer()
    .then(buffer => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'op.xlsx';
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
