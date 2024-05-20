<template>
  <div class="  w-full flex flex-col  ">
    <div class="truncate w-full flex-1">This is a secure webpage for token: {{ token }}</div>
    <div class="truncate  flex-1">This is a secure webpage for server: {{ server }}</div>
    <div class="truncate  flex-1">This is a response:
      <tr /> {{ response }}
    </div>
    <button class="w-16 bg-green-400 border-emerald-50 border-collapse border text-white"
      @click="exportToExcel">exportToExcel</button>
  </div>
</template>
<script setup>
import { ref, onMounted } from 'vue'
import { useRoute } from 'vue-router'
import axios from 'axios';

const route = useRoute()
const token = ref('')
const server = ref('')
const response = { data: {} }

onMounted(() => {
  token.value = route.query.token
  server.value = route.query.server
  const config = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token.value,
    },
    url: server.value + '/api/examReservations/export',
  };
  axios(config).then((res, err) => {
    response.data = res.data.data
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
      obj.Rt_dbHL_aided = dataList[4]
      obj.Lt_dis_aided = dataList[5]
      obj.Lt_dbHL_aided = dataList[6]
      break;
    case "ls_proun_score":
      obj.Rt_SRT = dataList[1]
      obj.Lt_SRT = dataList[2]
      obj.Rt_Dis = dataList[3]
      obj.Rt_dbHL = dataList[4]
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

// 주어진 데이터

function exportToExcel() {
  console.time('textColor')
  const workbook = new ExcelJS.Workbook();
  const workSheet = workbook.addWorksheet('인지청력검사');
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
    u_lang_test: "언어검사일",
    u_cog_test: "인지청력검사일",
    u_eeg_test: "EEG 검사일",
    // u_robot_test: "로봇 검사일",
    // u_result_category: "대상자 분류",
    // e_type:"e_type",
    e_date: "검사예정일",
    e_user_repeat: "회차",
    // isValid:"isValid",
    editBy: "데이터생성자",
    isDone: "완료여부",
    // createdAt: "생성일",
    // updatedAt: "업데이트일",
    m_name: "관리자",
    // ls_summary_id: "회차",
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
        { text: '(speech recognition threshold)_Aided', font: { size: 8 } },
      ]
    },
    Lt_SRT_aided: "Lt.-SRT(speech recognition threshold)_Aided",
    Rt_Dis_aided: "	Rt. Discrimination_Aided",
    Rt_dbHL_aided: "Rt. dbHL(m)_Aided",
    Lt_dis_aided: "Lt. discrimination_Aided",
    Lt_dbHL_aided: "Lt. dbHL(m)_Aided",
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
    rs_answer: " rs_answer",
    rs_question_num: "rs_question_num",
    rs_res_id: " rs_res_id",
    rs_summary_id: " rs_summary_id",
    rs_total_question_num: " rs_total_question_num",
    rs_type: " rs_type",
    rs_user_id: " rs_user_id",
  };
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
  setupTitleCell(workSheet, 'C1', 'A.기본정보', { alingleft: true, bold: true, merge: 'X1', border: borderStyle.side });
  setupTitleCell(workSheet, 'Y1', 'B.청력(순음 청력검사)', { alignRight: true, bold: true, merge: 'AV1' });
  setupTitleCell(workSheet, 'AW1', 'B.청력(어음청력검사)', { alignRight: true, bold: true, merge: 'BB1' });
  setupTitleCell(workSheet, 'BC1', 'B.청력(Aided)', { alignRight: true, bold: true, merge: 'BH1' });
  setupTitleCell(workSheet, 'BI1', 'B.청력(임피던스 검사)', { alignRight: true, bold: true, merge: 'BJ1' });
  setupTitleCell(workSheet, 'BK1', 'B.청력(보청기착용 우측)', { alignRight: true, bold: true, merge: 'BP1' });
  setupTitleCell(workSheet, 'BQ1', 'B.청력(보청기착용 좌측)', { alignRight: true, bold: true, merge: 'BV1' });
  setupTitleCell(workSheet, 'BW1', 'B.청력(난청유무)', { alignRight: true, bold: true, merge: 'CH1' });
  setupTitleCell(workSheet, 'CI1', 'B 청력(난청정도)', { alignRight: true, bold: true, merge: 'CJ1' });
  //*************************************************************/

  for (const item of response.data) {


    try {
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
      if (item.ls_proun_score != null) {
        addDataToObj(item.ls_proun_score, item, 'ls_proun_score')
      }
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
            item[key] = dayjs(item[key]).format("YYYY-MM-DD")
          }
      });
      // excel에 라인추가 
      workSheet.addRow(item)
    } catch (error) {
      console.error('error :>> ', error);
    }
  }
  const ls_key =
  {
    createdAt: "2024-05-14T04:15:39.000Z",
    editBy: "admin",
    ls_imp_score: ",-1,-1",
    ls_loss_score: ",false,false,false,false,false,false,-1,false,false,false,false,false,false,-1",
    // aided
    ls_proun_aided_score: ",1,-1,3,4,5,5",
    // 어음청력검사
    ls_proun_score: ",1,2,3,4,5,6",
    // 골도기도
    ls_user_id: "136",
    updatedAt: "2024-05-14T04:16:07.000Z",

    ls_aided_score:
      ",2,-1,-1,-1,-1,-1,-1,34,-1,5,-1,-1",

    ls_pure_score
      :
      ",1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,78,8,9,0,1,2,3,4",

    // ls_pure_score 데이터 순서별로 



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
    column.eachCell({ includeEmpty: true }, cell => {
      const cellValue = cell.value ? cell.value.toString() : '';
      // 너비를 픽셀 단위로 변환 (약간의 여유 공간 추가)
      const cellWidth = cellValue.length + 2;
      // 현재 열의 너비와 비교하여 큰 값을 사용
      if (!column.width || column.width < cellWidth) {
        column.width = cellWidth;
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
