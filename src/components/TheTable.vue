<template>
  <div class="sheet-table">
    <div ref="tableau" v-html="html"></div>
    <button @click="exportFile">Export XLSX</button>
  </div>
</template>

<script setup>
import { ref, onMounted } from "vue";
import { read, utils, writeFileXLSX } from "xlsx";

const html = ref("");
const tableau = ref();

onMounted(async () => {
  /* Download from https://docs.sheetjs.com/pres.numbers */
  const f = await fetch("https://docs.sheetjs.com/pres.numbers");
  const ab = await f.arrayBuffer();

  /* parse workbook */
  const wb = read(ab);

  /* update data */
  html.value = utils.sheet_to_html(wb.Sheets[wb.SheetNames[0]]);
});

/* get live table and export to XLSX */
function exportFile() {
  const wb = utils.table_to_book(
    tableau.value.getElementsByTagName("TABLE")[0]
  );
  writeFileXLSX(wb, "SheetJSVueHTML.xlsx");
}
</script>

<style lang="scss" scoped></style>
