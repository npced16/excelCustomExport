<template>
  <div class="w-full flex flex-col">
    <div v-if="!fileDownloadFlag" class="w-full h-[90vh] flex items-center justify-center text-center text-8xl">
      엑셀파일을 다운로드중입니다.
    </div>
    <div v-else class="w-full h-[90vh] flex flex-col items-center justify-center text-center text-8xl">
      엑셀파일을 완료했습니다.
      <div>
        <span class="text-red-500 text-10xl">{{ countTime }}</span> 후에 창이 종료됩니다.
      </div>
    </div>
  </div>

</template>
<script setup>
import { ref, onMounted } from 'vue'
import { useRoute } from 'vue-router'
import axios from 'axios';

const route = useRoute()
const token = ref('')
const server = ref('')
const autoClose = ref(false)
const response = { data: {} }
// window.close();
onMounted(() => {
  token.value = route.query.token
  server.value = route.query.server
  autoClose.value = route.query.autoClose
  const config = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token.value,
    },
    url: server.value + '/api/examReservations/export',
  };
  axios(config).then((res, err) => {
    // response.data = res.data.data
    exportToExcel(res.data.data)
    console.table(res.data.data);
  }).catch((err) => {
    console.error(err)
  });

})
import ExcelJS from 'exceljs';
import { backgroundColors as bgColors, textColor, defaultCellStyle, borderStyle } from "@/plugins/style"
import {
  setupTitleCell, toBoldText, setCellFill
  , setCellBorder, setCellColor, determineReservationsTitleStyles
} from "@/plugins/chart.js"
import dayjs from 'dayjs'
const fileDownloadFlag = ref(false)
const countTime = ref(5)

function addDataToObj(dataString, obj, version) {
  var dataList = dataString.split(',');
  // dataList의 값이 -1이면 '-'로 바꾸는 부분
  // if()
  // dataList = dataList.map(value => value === -1 ? "-" : value);

  switch (version) {
    case "rs_answer_cerad_answer":
      for (var i = 1; i <= 22; i++) {
        obj["rs_CERAD_" + i] = dataList[i * 2 - 1];
        obj["rs_CERAD_" + i + "_zScore"] = dataList[i * 2];
      }
      break;
    case "rs_answer_geriatric_answer":
      obj.rs_GDS_KR_true = dataList[1]
      obj.rs_GDS_KR_false = dataList[2]
      break;
    case "rs_answer_global_answer":
      obj.rs_GDS = dataList[1]
      break;
    case "rs_answer_iadl_answer":
      obj.rs_IADL_1 = dataList[1]
      obj.rs_IADL_2 = dataList[2]
      obj.rs_IADL_3 = dataList[3]
      obj.rs_IADL_4 = dataList[4]
      obj.rs_IADL_5 = dataList[5]
      obj.rs_IADL_6 = dataList[6]
      obj.rs_IADL_7 = dataList[7]
      obj.rs_IADL_8 = dataList[8]
      break;
    case "ls_pure_score":
      obj.r_bone_250hz = dataList[1]
      obj.r_bone_500hz = dataList[2]
      obj.r_bone_1Khz = dataList[3]
      obj.r_bone_2Khz = dataList[4]
      obj.r_bone_4Khz = dataList[5]
      obj.r_bone_8Khz = dataList[6]
      obj.r_air_250hz = dataList[7]
      obj.r_air_500hz = dataList[8]
      obj.r_air_1Khz = dataList[9]
      obj.r_air_2Khz = dataList[10]
      obj.r_air_4Khz = dataList[11]
      obj.r_air_8Khz = dataList[12]
      obj.l_bone_250hz = dataList[13]
      obj.l_bone_500hz = dataList[14]
      obj.l_bone_1Khz = dataList[15]
      obj.l_bone_2Khz = dataList[16]
      obj.l_bone_4Khz = dataList[17]
      obj.l_bone_8Khz = dataList[18]
      obj.l_air_250hz = dataList[19]
      obj.l_air_500hz = dataList[20]
      obj.l_air_1Khz = dataList[21]
      obj.l_air_2Khz = dataList[22]
      obj.l_air_4Khz = dataList[23]
      obj.l_air_8Khz = dataList[24]
      break;
    case "ls_proun_aided_score":
      obj.Rt_SRT_aided = dataList[1]
      obj.Lt_SRT_aided = dataList[2]
      obj.Rt_Dis_aided = dataList[3]
      obj.Rt_dbmL_aided = dataList[4]
      obj.Lt_dis_aided = dataList[5]
      obj.Lt_dbHL_aided = dataList[6]
      break;
    case "ls_proun_score":
      obj.Rt_SRT = dataList[1]
      obj.Lt_SRT = dataList[2]
      obj.Rt_Dis = dataList[3]
      obj.Rt_dbmL = dataList[4]
      obj.Lt_dis = dataList[5]
      obj.Lt_dbHL = dataList[6]
      break;
    case "ls_imp_score":
      obj.imp_rt = dataList[1]
      obj.imp_lt = dataList[2]
      break;
    case "ls_aided_score":
      obj.r_250hz_aided = dataList[1]
      obj.r_500hz_aided = dataList[2]
      obj.r_1Khz_aided = dataList[3]
      obj.r_2Khz_aided = dataList[4]
      obj.r_4Khz_aided = dataList[5]
      obj.r_8Khz_aided = dataList[6]
      obj.l_250hz_aided = dataList[7]
      obj.l_500hz_aided = dataList[8]
      obj.l_1Khz_aided = dataList[9]
      obj.l_2Khz_aided = dataList[10]
      obj.l_4Khz_aided = dataList[11]
      obj.l_8Khz_aided = dataList[12]
      break;
    case "ls_loss_score":
      obj.loss_r_250hz = dataList[1] !== "false" ? 0 : 'x'
      obj.loss_r_500hz = dataList[2] !== "false" ? 0 : 'x'
      obj.loss_r_1Khz = dataList[3] !== "false" ? 0 : 'x'
      obj.loss_r_2Khz = dataList[4] !== "false" ? 0 : 'x'
      obj.loss_r_4Khz = dataList[5] !== "false" ? 0 : 'x'
      obj.loss_r_8Khz = dataList[6] !== "false" ? 0 : 'x'
      obj.loss_r_pta = dataList[7]
      obj.loss_l_250hz = dataList[8] !== "false" ? 0 : 'x'
      obj.loss_l_500hz = dataList[9] !== "false" ? 0 : 'x'
      obj.loss_l_1Khz = dataList[10] !== "false" ? 0 : 'x'
      obj.loss_l_2Khz = dataList[11] !== "false" ? 0 : 'x'
      obj.loss_l_4Khz = dataList[12] !== "false" ? 0 : 'x'
      obj.loss_l_8Khz = dataList[13] !== "false" ? 0 : 'x'
      obj.loss_l_pta = dataList[14]
      break;
  }
  return obj;
}
// key 값은 바꿔도됌 
const workSheetKeyList = {
  e_id: "index",
  u_acc_code: "증례번호",
  // e_q_paper_ids:"e_q_paper_ids",
  u_id: "유저번호",
  u_name: "성명",
  u_sex: "성별",
  u_birth: "생년월일",
  u_telephone: "전화번호",
  u_agency_code: "기관코드",
  u_chart_number: "병록번호",
  u_enter_path: "유입 경로",
  u_study_year: "학력",
  u_blank: "비고",
  u_kbase_test: "KBASE2",
  u_kbase_move_date: "KBASE2 이관일",
  u_kbase_result_date: "KBASE2 최종통보일",
  u_kbase_result: "KBASE2 결과",
  u_lang_test: "언어 검사일",
  u_cog_test: "인지청력 검사일",
  u_eeg_test: "EEG 검사일",
  // u_robot_test: "로봇 검사일",
  // u_result_category: "대상자 분류",
  // e_type:"e_type",
  e_date: "검사 예정일",
  e_user_repeat: "회차",
  // isValid:"isValid",
  // editBy: "데이터생성자",
  // isDone: "완료여부",
  // createdAt: "생성일",
  // updatedAt: "업데이트일",
  // m_name: "관리자",
  rs_CERAD_2: "J1. 언어유창성 검사 : 동물 범주",
  rs_CERAD_3: "J2. 보스톤 이름대기 검사",
  rs_CERAD_1: "J3. MMSE-KC",
  rs_CERAD_22: "J3. MMSE-KC(정신과에서 측정)",
  rs_CERAD_4: "J4. 단어목록 기억검사",
  rs_CERAD_5: "J4. 시행(1)",
  rs_CERAD_6: "J4. 시행(2)",
  rs_CERAD_7: "J4. 시행(3)",
  rs_CERAD_8: "J5. 구성행동 검사",
  rs_CERAD_9: "J6. 단어목록 회상검사",
  rs_CERAD_10: "J6.회상률(%)",
  rs_CERAD_11: "J7. 단어목록 재인검사",
  rs_CERAD_12: "예",
  rs_CERAD_13: "아니오",
  rs_CERAD_14: "J8. 구성회상 검사",
  rs_CERAD_15: "길찾기 A",
  rs_CERAD_16: "길찾기 B",
  rs_CERAD_17: "J가.단어페이지",
  rs_CERAD_18: "J가.색깔페이지",
  rs_CERAD_19: "J가.색깔 - 단어 페이지",
  rs_CERAD_20: "총점 I(J1+J2+J4+J5+J6+J7)",
  rs_CERAD_21: "총점 II(J1+J2+J4+J5+J6+J7+J8)",
  rs_CERAD_2_zScore: "J1.z-score",
  rs_CERAD_3_zScore: "J2.z-score",
  rs_CERAD_1_zScore: "J3.z-score",
  rs_CERAD_22_zScore: "J3.z-score",
  rs_CERAD_4_zScore: "J4.z-score",
  rs_CERAD_5_zScore: "J4(1).z-score",
  rs_CERAD_6_zScore: "J4(2).z-score",
  rs_CERAD_7_zScore: "J4(3).z-score",
  rs_CERAD_8_zScore: "J5.z-score",
  rs_CERAD_9_zScore: "J6.z-score",
  rs_CERAD_10_zScore: "J6(회상률).z-score",
  rs_CERAD_11_zScore: "J7(통합).z-score",
  rs_CERAD_12_zScore: "J7(예).z-score",
  rs_CERAD_13_zScore: "J7(아니오).z-score",
  rs_CERAD_14_zScore: "J8.z-score",
  rs_CERAD_15_zScore: "J9(A).z-score",
  rs_CERAD_16_zScore: "J9(B).z-score",
  rs_CERAD_17_zScore: "J가(단어).z-score",
  rs_CERAD_18_zScore: "J가(색깔).z-score",
  rs_CERAD_19_zScore: "J가(색깔-단어).z-score",
  rs_CERAD_20_zScore: "총점 I.z-score",
  rs_CERAD_21_zScore: "총점 II.z-score",
  //  ls_summary_id: "회차",
  rs_IADL_1: "A. 전화 사용능력",
  rs_IADL_2: "B. 물건사기",
  rs_IADL_3: "C. 음식준비하기",
  rs_IADL_4: "D. 집안일 하기",
  rs_IADL_5: "E. 빨래하기",
  rs_IADL_6: "F. 교통수단 이용",
  rs_IADL_7: "G. 약 복용하기",
  rs_IADL_8: "H. 돈 관리 능력",
  rs_GDS_KR_true: "GDS-KR(Geriatric Depression Scale)\n 긍정점수",
  rs_GDS_KR_false: "GDS-KR(Geriatric Depression Scale)\n 부정점수",
  rs_GDS: "GDS(Global Deterioration scale)",
  // BQ1: "GDS(Global Deterioration scale)",
  r_bone_250hz: "(우)골도 250hz",
  r_bone_500hz: "(우)골도 500hz",
  r_bone_1Khz: "(우)골도 1Khz",
  r_bone_2Khz: "(우)골도 2Khz",
  r_bone_4Khz: "(우)골도 4Khz",
  r_bone_8Khz: "(우)골도 8Khz",
  r_air_250hz: "(우)기도 250hz",
  r_air_500hz: "(우)기도 500hz",
  r_air_1Khz: "(우)기도 1Khz",
  r_air_2Khz: "(우)기도 2Khz",
  r_air_4Khz: "(우)기도 4Khz",
  r_air_8Khz: "(우)기도 8Khz",
  l_bone_250hz: "(좌)골도 250hz",
  l_bone_500hz: "(좌)골도 500hz",
  l_bone_1Khz: "(좌)골도 1Khz",
  l_bone_2Khz: "(좌)골도 2Khz",
  l_bone_4Khz: "(좌)골도 4Khz",
  l_bone_8Khz: "(좌)골도 8Khz",
  l_air_250hz: "(좌)기도 250hz",
  l_air_500hz: "(좌)기도 500hz",
  l_air_1Khz: "(좌)기도 1Khz",
  l_air_2Khz: "(좌)기도 2Khz",
  l_air_4Khz: "(좌)기도 4Khz",
  l_air_8Khz: "(좌)기도 8Khz",
  Rt_SRT: "Rt.-SRT(speech recognition threshold)",
  Lt_SRT: "Lt.-SRT(speech recognition threshold)",
  Rt_Dis: "	Rt. Discrimination",
  Rt_dbHL: "Rt. dbHL(m)",
  Lt_dis: "Lt. discrimination",
  Lt_dbHL: "Lt. dbHL(m)",
  Rt_SRT_aided:
  {
    richText: [
      { text: 'Rt.-SRT\n', font: { bold: true, } },
      { text: '(speech recognition threshold)', font: { size: 8 } },
    ]
  },
  Lt_SRT_aided: "Lt.-SRT(speech recognition threshold)",
  Rt_Dis_aided: "	Rt. Discrimination",
  Rt_dbmL_aided: "Rt. dbHL(m)",
  Lt_dis_aided: "Lt. discrimination",
  Lt_dbHL_aided: "Lt. dbHL(m)",
  imp_rt: "임피던스_Rt. Side",
  imp_lt: "임피던스_Lt. Side",
  r_250hz_aided: "우측 250hz",
  r_500hz_aided: "우측 500hz",
  r_1Khz_aided: "우측 1Khz",
  r_2Khz_aided: "우측 2Khz",
  r_4Khz_aided: "우측 4Khz",
  r_8Khz_aided: "우측 8Khz",
  l_250hz_aided: "좌측 250hz",
  l_500hz_aided: "좌측 500hz",
  l_1Khz_aided: "좌측 1Khz",
  l_2Khz_aided: "좌측 2Khz",
  l_4Khz_aided: "좌측 4Khz",
  l_8Khz_aided: "좌측 8Khz",
  loss_r_250hz: "우측 250hz",
  loss_r_500hz: "우측 500hz",
  loss_r_1Khz: "우측 1Khz",
  loss_r_2Khz: "우측 2Khz",
  loss_r_4Khz: "우측 4Khz",
  loss_r_8Khz: "우측 8Khz",
  loss_l_250hz: "좌측 250hz",
  loss_l_500hz: "좌측 500hz",
  loss_l_1Khz: "좌측 1Khz",
  loss_l_2Khz: "좌측 2Khz",
  loss_l_4Khz: "좌측 4Khz",
  loss_l_8Khz: "좌측 8Khz",
  loss_r_pta: "우측(PTA평균)",
  loss_l_pta: "좌측(PTA평균)",
};
// 주어진 데이터

function exportToExcel(rawData) {
  console.time('textColor')
  const workbook = new ExcelJS.Workbook();
  const workSheet = workbook.addWorksheet('인지청력검사');

  workSheet.columns = Object.keys(workSheetKeyList).map((key) => (
    {
      header: '',
      key: key,
      style: (key == 'C1') ? { ...defaultCellStyle, font: { bold: true } } : defaultCellStyle,
    })
  );

  workSheet.addRow(workSheetKeyList)

  //헤더설정 /
  // setupTitleCell(workSheet, 'C1', '',);
  setupTitleCell(workSheet, 'C1', 'A.기본정보', { alingleft: true, bold: true, merge: 'U1', border: borderStyle.side });
  // setupTitleCell(workSheet, 'Y1', 'B.인지-CERAD-K', { alignRight: true, bold: true, merge: 'BP1' });
  // setupTitleCell(workSheet, 'BQ1', 'B.인지-IADL(Instrumental Activities of Daily Living)', { alignRight: true, bold: true, merge: 'BX1' });
  // setupTitleCell(workSheet, 'BY1', 'B.인지-GDS-KR(Geriatric Depression Scale)', { alignleft: true, bold: true, merge: 'BZ1' });
  // setupTitleCell(workSheet, 'CA1', 'B.인지-GDS(Global Deterioration scale) ', { bold: true });
  // setupTitleCell(workSheet, 'CB1', 'C.청력(순음 청력검사)', { alignRight: true, bold: true, merge: 'CY1' });
  // setupTitleCell(workSheet, 'DA1', 'C.청력(어음청력검사)', { alignRight: true, bold: true, merge: 'DE1' });
  // setupTitleCell(workSheet, 'DF1', 'C.청력(보청기착용)', { alignRight: true, bold: true, merge: 'DK1' });
  // setupTitleCell(workSheet, 'DL1', 'C.청력(임피던스 검사)', { alignRight: true, bold: true, merge: 'DM1' });
  // setupTitleCell(workSheet, 'DN1', 'C.청력(보청기착용 우측)', { alignRight: true, bold: true, merge: 'DS1' });
  // setupTitleCell(workSheet, 'DT1', 'C.청력(보청기착용 좌측)', { alignRight: true, bold: true, merge: 'DY1' });
  // setupTitleCell(workSheet, 'DZ1', 'C.청력(난청유무)', { alignRight: true, bold: true, merge: 'EK1' });
  // setupTitleCell(workSheet, 'EL1', 'C 청력(난청정도)', { alignRight: true, bold: true, merge: 'EM1' });

  setupTitleCell(workSheet, 'X1', 'B.인지-CERAD-K', { alignRight: true, bold: true, merge: 'AN1' });
  setupTitleCell(workSheet, 'BN1', 'B.인지-IADL(Instrumental Activities of Daily Living)', { alignRight: true, bold: true, merge: 'BR1' });
  setupTitleCell(workSheet, 'BR1', 'B.인지-GDS-KR(Geriatric Depression Scale)', { alignLeft: true, bold: true, merge: 'BS1' });
  setupTitleCell(workSheet, 'BS1', 'B.인지-GDS(Global Deterioration scale) ', { bold: true });
  setupTitleCell(workSheet, 'BT1', 'C.청력(순음 청력검사)', { alignRight: true, bold: true, merge: 'BS1' });
  setupTitleCell(workSheet, 'BR1', 'C.청력(어음청력검사)', { alignRight: true, bold: true, merge: 'DC1' });
  setupTitleCell(workSheet, 'DC1', 'C.청력(보청기착용)', { alignRight: true, bold: true, merge: 'DF1' });
  setupTitleCell(workSheet, 'DF1', 'C.청력(임피던스 검사)', { alignRight: true, bold: true, merge: 'DG1' });
  setupTitleCell(workSheet, 'DH1', 'C.청력(보청기착용 우측)', { alignRight: true, bold: true, merge: 'DP1' });
  setupTitleCell(workSheet, 'DP1', 'C.청력(보청기착용 좌측)', { alignRight: true, bold: true, merge: 'DV1' });
  setupTitleCell(workSheet, 'DV1', 'C.청력(난청유무)', { alignRight: true, bold: true, merge: 'EA1' });
  setupTitleCell(workSheet, 'EA1', 'C 청력(난청정도)', { alignRight: true, bold: true, merge: 'EB1' });



  //*************************************************************/



  for (const item of rawData) {


    try {
      if (item.rs_answer_cerad_answer != null) {
        addDataToObj(item.rs_answer_cerad_answer, item, 'rs_answer_cerad_answer')
      }
      if (item.rs_answer_iadl_answer != null) {
        addDataToObj(item.rs_answer_iadl_answer, item, 'rs_answer_iadl_answer')
      }
      if (item.rs_answer_geriatric_answer != null) {
        addDataToObj(item.rs_answer_geriatric_answer, item, 'rs_answer_geriatric_answer')
      }
      if (item.rs_answer_global_answer != null) {
        addDataToObj(item.rs_answer_global_answer, item, 'rs_answer_global_answer')
      }
      // 청력검사내용 바인딩
      if (item.ls_pure_score != null) {
        addDataToObj(item.ls_pure_score, item, 'ls_pure_score')
      }
      // 어음청력검사 )보청기
      if (item.ls_proun_aided_score != null) {
        addDataToObj(item.ls_proun_aided_score, item, 'ls_proun_aided_score')
      }
      //보청기
      if (item.ls_aided_score != null) {
        addDataToObj(item.ls_aided_score, item, 'ls_aided_score')
      }
      // 인피던스
      if (item.ls_imp_score != null) {
        addDataToObj(item.ls_imp_score, item, 'ls_imp_score')
      }
      // 
      if (item.ls_proun_score != null) {
        addDataToObj(item.ls_proun_score, item, 'ls_proun_score')
      }
      // 
      if (item.ls_loss_score != null) {
        addDataToObj(item.ls_loss_score, item, 'ls_loss_score')
      }




      // 빈값 채워넣기 및 포맷팅 
      Object.keys(workSheetKeyList).forEach(key => {
        if (item[key] == null) {
          item[key] = '-';
        } else
          if (key == "createdAt" || key == "updatedAt" || key == "u_lang_test" || key == "e_date"
            || key == "u_kbase_move_date" || key == "u_kbase_result_date"
            || key == "u_lang_test" || key == "u_cog_test" || key == "u_eeg_test"
          ) {
            // item[key] = dayjs(item[key]).format("YYYY-MM-DD")
            if (item[key] !== '-') {
              item[key] = dayjs(item[key]).format("YYYY-MM-DD");
            }
          }
      });

      // excel에 라인추가 
      workSheet.addRow(item)
    } catch (error) {
      console.error('error :>> ', error);
    }
  }
  const headerNumber = 2
  workSheet.eachRow({ includeEmpty: true },
    function (row, rowNumber) {
      row.eachCell(function (cell, colNumber) {
        // title 색상처리등 
        if (rowNumber === headerNumber) {
          row.height = 80; // 높이가 60 /1.5 -> 40 나옴 
          // row.vertical = 120;
          toBoldText(cell)
          const { bgColor, borderType } = determineReservationsTitleStyles(colNumber);
          setCellFill(cell, bgColor);
          setCellBorder(cell, borderType)
          setCellLineBreak(cell)
        }
      });
    }
  )
  // 각 셀의 내용에 따라 열 너비를 설정
  workSheet.columns.forEach(column => {
    column.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
      if (rowNumber !== 1) {
        const cellValue = cell.value ? cell.value.toString() : '';
        const cellWidth = cellValue.length + 2;
        if (!column.width || column.width < cellWidth) {
          column.width = cellWidth;
        }
      }
    });
  });

  addFileter(workSheet, headerNumber)

  // addFileter(workSheet, headerNumber)
  // 파일 다운로드
  downloadxlsx(workbook, '청력')
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
  // window.close()
}
function countDown() {
  fileDownloadFlag.value = true
  // if (autoClose.value == true) {
  setInterval(() => {
    countTime.value--; // 카운트다운 값을 1씩 줄임
    if (countTime.value === 0) {
      window.close(); // 창을 닫음
    }
  }, 1000); // 1초마다 실행
  // }

}

function setCellLineBreak(cell) {
  cell.style = {
    ...cell.style,
    alignment: { wrapText: true, horizontal: 'center', vertical: 'middle' }
  }
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

</script>

<style scoped></style>
