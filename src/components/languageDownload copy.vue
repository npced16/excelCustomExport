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
  , setCellBorder, setCellColor, determineReservationsTitleStyles
} from "@/plugins/chart.js"
const fileDownloadFlag = ref(false)
const countTime = ref(5)


function exportToExcel() {
  console.time('textColor')
  const workbook = new ExcelJS.Workbook();
  const workSheet = workbook.addWorksheet('인지청력검사');
  // key 값은 바꿔도됌 
  const workSheetKeyList = {
    e_id: "유저번호",
    e_user_id: "유저아이디",
    e_q_paper_ids: "e_q_paper_ids",
    u_id: "아이디",
    u_name: "이름",
    u_sex: "성별",
    u_birth: "생년월일",
    u_telephone: "전화번호",
    u_agency_code: "u_agency_code",
    u_acc_code: "증례번호",
    u_chart_number: "u_chart_number",
    u_enter_path: "u_enter_path",
    u_study_year: "u_study_year",
    u_blank: "u_blank",
    u_cog_test: "u_cog_test",
    u_kbase_test: "u_kbase_test",
    u_kbase_move_date: "u_kbase_move_date",
    u_kbase_result_date: "u_kbase_result_date",
    u_lang_test: "u_lang_test",
    u_eeg_test: "u_eeg_test",
    eeg_id: "eeg_id",
    e_type: "e_type",
    e_date: "e_date",
    e_user_repeat: "e_user_repeat",
    isValid: "isValid",
    editBy: "editBy",
    isDone: "isDone",
    createdAt: "createdAt",
    updatedAt: "updatedAt",
    m_name: "m_name",
  };
  workSheet.columns = Object.keys(workSheetKeyList).map((key) => (
    {
      header: '',
      key: key,
      style: (key == 'C1') ? { ...defaultCellStyle, font: { bold: true } } : defaultCellStyle,
    })
  );

  workSheet.addRow(workSheetKeyList)

  setupTitleCell(workSheet, 'C1', '', { border: borderStyle.left });
  setupTitleCell(workSheet, 'D1', 'A.기본정보', { bold: true, merge: 'E1' });
  console.table(response.data);
  for (const item of response.data) {
    workSheet.addRow(item)
  }
  const dataLength = response.data.length

  // 데이터 길이만큼
  const lastRownumber = dataLength + 2
  const headerNumber = 2



  workSheet.eachRow({ includeEmpty: true },
    function (row, rowNumber) {
      row.eachCell(function (cell, colNumber) {
        // title 색상처리등 
        if (rowNumber === headerNumber) {
          row.height = 80; // 높이가 60 /1.5 -> 40 나옴 
          row.vertical = 120;
          toBoldText(cell)
          const { bgColor, borderType } = determineReservationsTitleStyles(colNumber);
          setCellFill(cell, bgColor);
          setCellBorder(cell, borderType)
          setCellLineBreak(cell)
        }
      });
    }
  )
  // addFileter(workSheet, headerNumber)
  // 파일 다운로드
  downloadxlsx(workbook, 'aaaRandom')
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
