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
    url: server.value + '/api/examUsers',
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
  , setCellBorder, setCellColor, determineTitleStyles
} from "@/plugins/chart.js"
import dayjs from 'dayjs'
const fileDownloadFlag = ref(false)
const countTime = ref(5)
const de =
{
  createdAt: "2024-05-14T04:15:39.000Z",
  editBy: "admin",
  ls_aided_score: ",-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1",
  ls_imp_score: ",-1,-1",
  ls_loss_score: ",false,false,false,false,false,false,-1,false,false,false,false,false,false,-1",
  ls_proun_aided_score: ",1,-1,3,4,5,5",
  ls_proun_score: ",1,2,3,4,5,6",
  ls_pure_score: ",1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,78,8,9,0,1,2,3,4",
  ls_res_id: 143,
  ls_summary_id: 1,
  ls_user_id: "136",
  updatedAt: "2024-05-14T04:16:07.000Z",
  // ls_pure_score 데이터 순서별로 
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
  l_air_8Khz: "(좌)기도 8Khz"

}

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
    u_enter_path: "유입",
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
    createdAt: "생성일",
    updatedAt: "업데이트일",
    m_name: "관리자",
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
  setupTitleCell(workSheet, 'C1', 'A.기본정보', { alingleft: true, bold: true, merge: 'Z1', border: borderStyle.side });

  //*************************************************************/


  console.table(response.data);
  for (const item of response.data) {
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
    workSheet.addRow(item)
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
          const { bgColor, borderType } = determineTitleStyles(colNumber);
          setCellFill(cell, bgColor);
          setCellBorder(cell, borderType)
          setCellLineBreak(cell)
        }

      });
    }
  )
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
