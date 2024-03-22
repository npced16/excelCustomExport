<template>
  <button @click="exportToExcel">Export to Excel</button>
  <div v-if="!fileDownloadFlag" style="width: 100%; height: 90vh;
    display: flex; text-align: center; align-items: center; justify-content: center; font-size: 8rem;">
    엑셀파일을 다운로드중입니다.
  </div>
  <div v-else
    style="width: 100%; height: 90vh;
    display: flex; flex-direction: column; text-align: center; align-items: center; justify-content: center; font-size: 8rem;">
    엑셀파일을 완료했습니다.
    <div>

      <span style="color: red; font-size: 10rem; ">{{ countTime }}</span> 후에 창이 종료됩니다.
    </div>
  </div>
</template>

<script setup>
import ExcelJS from 'exceljs';
import { backgroundColors as bgColors, textColor, defaultCellStyle, borderStyle } from "@/plugins/style"
import { ref } from 'vue';
const fileDownloadFlag = ref(false)
const countTime = ref(5)
/**
 * Header 셀 설정 
 * @param {*} ws  workSheet
 * @param {*} cellRef cellr 번호 ex) A1,QO2
 * @param {*} value  표시할 내용
 * @param {*} options  bold , border , alignRight , merge
 */
function setupTitleCell(ws, cellRef, value, options = {}) {
  const cell = ws.getCell(cellRef);
  cell.value = value;
  if (options.bold) {
    toBoldText(cell)
  }
  if (options.border) {
    setCellBorder(cell, options?.border)
  }
  if (options.alignRight) {
    cell.alignment = { horizontal: 'right' };
  }
  if (options.merge) {
    ws.mergeCells(cellRef + ':' + options.merge);
  }
}

// text bold 하게 만드는 함수 (기존 폰트 유지하며)
function toBoldText(cell) {
  cell.font = { ...cell.font, bold: true }
}
// 지정된 색으로 text 칠하는 함수
function setCellFill(cell, color) {
  if (color) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: color },
    };
  }
}
// 주어진 style 로 cell의 border 설정 
function setCellBorder(cell, style) {
  if (style) {
    cell.border = { ...style.border };
  }
}
//  change text color  
function setCellColor(cell, color) {
  cell.font = { ...(cell.font), color: { argb: color } };
}

// title 첫줄 의 
function determineTitleStyles(colNumber) {
  // title 연결하고 boerder 설정 하는 곳
  const styleConditions = [
    { range: [1, 2], bgColor: bgColors.basic, borderType: borderStyle.right_side },
    {
      range: [3, 13], bgColor: bgColors.start,
      borderType:
        (col) => (col === 11)
          ? borderStyle.right_side
          : borderStyle.top_bottom
    },
    {
      range: [57, 64], bgColor: bgColors.IADL,
      borderType: (col) =>
        (col === 57)
          ? borderStyle.left_side
          : ((col === 64)
            ? borderStyle.right_side
            : borderStyle.top_bottom)
    },
    { range: [65, 69], bgColor: bgColors.GDS, borderType: (col) => (col === 69) ? borderStyle.right_side : borderStyle.top_bottom },
    { range: [70, 93], bgColor: bgColors.a2, borderType: null },
    { range: [94, 99], bgColor: bgColors.SRT, borderType: null },
    { range: [100, 105], bgColor: bgColors.SRT2, borderType: null },
    { range: [106, 107], bgColor: bgColors.LT, borderType: null },
    { range: [108, 113], bgColor: bgColors.bo, borderType: null },
    { range: [120, 120], bgColor: null, borderType: null },
    { range: [121, 122], bgColor: bgColors.pta_avr, borderType: null },
    { range: [114, 119], bgColor: bgColors.bo_left, borderType: null },
  ];
  for (const condition of styleConditions) {
    const [start, end] = condition.range;
    if (colNumber >= start && colNumber <= end) {
      return {
        bgColor:
          typeof condition.bgColor === 'function'
            ? condition.bgColor(colNumber)
            : condition.bgColor,
        borderType:
          typeof condition.borderType === 'function'
            ? condition.borderType(colNumber)
            : condition.borderType,
      };
    }
  }
  const defaultStyle = { bgColor: bgColors.default, borderType: borderStyle.bottom };
  return defaultStyle;
}


function exportToExcel() {
  console.time('textColor')
  const workbook = new ExcelJS.Workbook();
  const workSheet = workbook.addWorksheet('인지청력검사');
  const ws1Header = {
    A1: {
      richText: [
        { text: '증례번호\n', font: { bold: true, } },
        { text: '(IRB 증례기록지용)', font: { size: 8 } },
      ]
    },
    B1: "연구대상자 이니셜\n(IRB 증례기록지용)",
    C1: "NO",
    D1: "대상자 번호",
    E1: "성별",
    F1: "검사일자",
    G1: "나이",
    H1: "교육년수",
    I1: "유입경로",
    J1: "기타",
    K1: {
      richText: [
        { text: '추후 참여 여부\n', font: { bold: true, } },
        { text: '*다른 연구 및 2~5년도 연구 계속 참여 여부', font: { size: 8 } },
      ]
    },
    L1: "종단 연구 의견",
    M1: "난청유무양이\n(PTA평균)",
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
  workSheet.columns = Object.keys(ws1Header).map((key) => (
    {
      header: '',
      key: key,
      style: (key == 'C1') ? { ...defaultCellStyle, font: { bold: true } } : defaultCellStyle,
    })
  );

  workSheet.addRow(ws1Header)


  setupTitleCell(workSheet, 'C1', '', { border: borderStyle.left });
  setupTitleCell(workSheet, 'D1', 'A.기본정보', { bold: true, merge: 'E1' });
  setupTitleCell(workSheet, 'M1', '', { border: borderStyle.right });
  setupTitleCell(workSheet, 'BE1', 'B.인지-IADL(Instrumental Activities of Daily Living)', { bold: true, merge: 'BL1' });
  setupTitleCell(workSheet, 'BM1', 'B.인지-GDS-KR(Geriatric Depression Scale)', { bold: true, merge: 'BP1' });
  setupTitleCell(workSheet, 'BQ1', 'B.인지-GDS(Global Deterioration scale)', { bold: true, border: borderStyle.right });
  setupTitleCell(workSheet, 'BR1', 'D.청력(순음 청력검사)', { alignRight: true, merge: 'CO1' });
  setupTitleCell(workSheet, 'CP1', 'D.청력(어음청력검사)', { alignRight: true, merge: 'CU1' });
  setupTitleCell(workSheet, 'CV1', 'D.청력(Aided)', { alignRight: true, merge: 'DA1' });
  setupTitleCell(workSheet, 'DB1', 'D.청력(임피던스 검사)', { merge: 'DC1' });
  setupTitleCell(workSheet, 'DD1', 'D.청력(보청기착용 우측)', { alignRight: true, merge: 'DI1' });
  setupTitleCell(workSheet, 'DJ1', 'D.청력(보청기착용 좌측)', { alignRight: true, merge: 'DO1' });
  setupTitleCell(workSheet, 'DP1', 'D.청력(난청유무)', { bold: true });
  setupTitleCell(workSheet, 'DQ1', 'D.청력(난청정도)', { merge: 'DR1' });

  const dataLength = getRandomDataLength(workSheet, ws1Header)
  // 데이터 길이만큼
  const lastRownumber = dataLength + 2
  const headerNumber = 2

  // 함수추가
  workSheet.fillFormula(`F3:F${lastRownumber}`, 'SUM(G3:J3)');
  // 
  workSheet.addConditionalFormatting({
    ref: `F3:F${lastRownumber}`,
    rules: [
      {
        type: 'cellIs',
        formulae: ['=0'],
        operator: "lessThan",
        style: { font: { color: { argb: textColor.red } } },
      }
    ]
  })

  workSheet.eachRow({ includeEmpty: true },
    function (row, rowNumber) {
      row.eachCell(function (cell, colNumber) {
        // title 색상처리등 
        if (rowNumber === headerNumber) {
          row.height = 60; // 높이가 60 /1.5 -> 40 나옴 
          toBoldText(cell)
          const { bgColor, borderType } = determineTitleStyles(colNumber);
          setCellFill(cell, bgColor);
          setCellBorder(cell, borderType)
        }
        // cell 색상처리등 
        if (rowNumber != headerNumber) {
          if (colNumber === 6 && cell.value < 0) {
            setCellColor(cell, textColor.red);
          }
        }
      });
    }
  )

  addFileter(workSheet, headerNumber)

  // 파일 다운로드
  downloadxlsx(workbook, 'Random')
}

async function downloadxlsx(workBook, fileName) {
  await workBook.xlsx.writeBuffer()
    .then(buffer => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = `${fileName}.xlsx`;
      link.click();
    })
    .catch(error => {
      console.error('Error exporting to Excel:', error);
    })
  countDown()
  console.timeEnd('textColor')
}
function countDown() {
  fileDownloadFlag.value = true
  setInterval(() => {
    countTime.value--; // 카운트다운 값을 1씩 줄임
    if (countTime.value === 0) {
      // window.close(); // 창을 닫음
    }
  }, 1000); // 1초마다 실행
}
function getRandomDataLength(workSheet, header) {
  //  랜덤값 집어넣기 
  let count = 1
  function getRandomValue(key) {
    if (key === "B1") {
      const randomAlphabet = () => String.fromCharCode(Math.floor(Math.random() * 26) + 65);
      return randomAlphabet() + randomAlphabet() + randomAlphabet();
    }
    if (key === "sumJ2_4") {
      return (Math.random() * 2) - 1;
    }
    if (key === "sex") {
      return Math.floor(Math.random() * 2) + 1;
    }
    if (key === " 시행1") {
      return (Math.random() * 2) - 1 > 0.5 ? (Math.random() * 2) - 1 : '';
    }
    if (key === "C1") {
      return count++;
    }
    return (Math.random() * 2) - 1 > -0.9 ? (Math.random() * 2) - 1 : "NA";
  }
  const randomObjects = Array.from({ length: 1000 }, () =>
    Object.fromEntries(Object.entries(header).map(([key]) => [key, getRandomValue(key)]))
  );
  for (const item of randomObjects) { workSheet.addRow(item) }
  return randomObjects.length
}

function addFileter(workSheet, headerNumber) {
  workSheet.autoFilter = {
    from: {
      row: headerNumber,
      column: 1
    },
    to: {
      row: headerNumber,
      column: workSheet.columns.length
    }
  };
}
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


</script>
