<template>
  <div class="sheet-table">
    <table>
      <thead class="sheet-table__header">
        <tr>
          <th>1</th>
          <th>Бренд</th>
          <th>Серия</th>
          <th>Осн.применение</th>
          <th>Цвет</th>
          <th>Основа</th>
          <th>Плотность</th>
          <th>Особенности</th>
          <th>Тип зерна</th>
          <th>Зернистость</th>

          <th class="abraforce">Abraforce</th>
          <th class="abraforce">P20</th>
          <th class="abraforce">P24</th>
          <th class="abraforce">P36</th>
          <th class="abraforce">P40</th>
          <th class="abraforce">P50</th>
          <th class="abraforce">P60</th>
          <th class="abraforce">P80</th>
          <th class="abraforce">P100</th>

          <th class="abraforce">P120</th>
          <th class="abraforce">P150</th>
          <th class="abraforce">P180</th>
          <th class="abraforce">P220</th>
          <th class="abraforce">P240</th>
          <th class="abraforce">P280</th>
          <th class="abraforce">P320</th>
          <th class="abraforce">P360</th>
          <th class="abraforce">P400</th>
          <th class="abraforce">P500</th>
          <th class="abraforce">P600</th>
          <th class="abraforce">P800</th>

          <th class="abraforce">P1000</th>
          <th class="abraforce">P1200</th>
          <th class="abraforce">P1500</th>
          <th class="abraforce">P2000</th>
          <th class="abraforce">P2500</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(row, idx) in rows" :key="idx">
          <td>{{ idx + 2 }}</td>
          <td>{{ row["Бренд"] }}</td>
          <td>{{ row["Серия"] }}</td>
          <td>{{ row["Осн.применение"] }}</td>
          <td>{{ row["Цвет"] }}</td>
          <td>{{ row["Основа"] }}</td>
          <td>{{ row["Плотность"] }}</td>
          <td>{{ row["Особенности"] }}</td>
          <td>{{ row["Тип зерна"] }}</td>
          <td>{{ row["Зернистость"] }}</td>
          <td>{{ row["Abraforce"] }}</td>

          <td>{{ row["P20"] }}</td>
          <td>{{ row["P24"] }}</td>
          <td>{{ row["P36"] }}</td>
          <td>{{ row["P40"] }}</td>
          <td>{{ row["P50"] }}</td>
          <td>{{ row["P60"] }}</td>
          <td>{{ row["P80"] }}</td>
          <td>{{ row["P100"] }}</td>

          <td>{{ row["P120"] }}</td>
          <td>{{ row["P150"] }}</td>
          <td>{{ row["P180"] }}</td>
          <td>{{ row["P220"] }}</td>
          <td>{{ row["P240"] }}</td>
          <td>{{ row["P280"] }}</td>
          <td>{{ row["P320"] }}</td>
          <td>{{ row["P360"] }}</td>
          <td>{{ row["P400"] }}</td>
          <td>{{ row["P500"] }}</td>
          <td>{{ row["P600"] }}</td>
          <td>{{ row["P800"] }}</td>

          <td>{{ row["P1000"] }}</td>
          <td>{{ row["P1200"] }}</td>
          <td>{{ row["P1500"] }}</td>
          <td>{{ row["P2000"] }}</td>
          <td>{{ row["P2500"] }}</td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<script setup>
import { ref, onMounted } from "vue";
import { read, utils } from "xlsx";

const table = ref();
const rows = ref([]);

const sheetToJson = async () => {
  const f = await fetch("/src/assets/table.xlsx");
  const ab = await f.arrayBuffer();

  /* parse workbook */
  const workbook = read(ab);

  /* JSON data */
  rows.value = utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

  console.log(rows.value);

  //raw_data[0]['Бренд'] = 'Новое название'

  /* update data */
  /* html.value = utils.sheet_to_html(utils.json_to_sheet(raw_data)); */

  //console.log(raw_data);
};

onMounted(async () => {
  sheetToJson();
});
</script>

<style lang="scss">
.sheet-table {
  $that: &;

  table {
    border-collapse: collapse;

    td {
      border: 1px solid rgb(148, 148, 148);
      padding: 2px 4px;
      text-align: center;
      vertical-align: middle;
    }
  }

  &__header {
    font-weight: bold;
    position: sticky;
    top: 0;
    background-color: rgb(207, 207, 207);

    th {
      cursor: pointer;
      padding: 6px 4px;
      border: 1px solid rgb(148, 148, 148);

      &.abraforce {
        background-color: rgb(177, 177, 255);
      }
    }
  }
}
</style>
