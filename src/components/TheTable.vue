<template>
  <div ref="table" class="sheet-table">
    <table>
      <thead class="sheet-table__header">
        <tr>
          <th>1</th>
          <th @click="sort(json, 'Бренд')">
            <span>Бренд <SortIcon /></span>
          </th>
          <th @click="sort(json, 'Серия')">
            <span>Серия <SortIcon /></span>
          </th>
          <th @click="sort(json, 'Осн.применение')">
            <span>Осн.применение <SortIcon /></span>
          </th>
          <th @click="sort(json, 'Цвет')">
            <span>Цвет <SortIcon /></span>
          </th>
          <th @click="sort(json, 'Основа')">
            <span>Основа <SortIcon /></span>
          </th>
          <th @click="sort(json, 'Плотность')">
            <span>Плотность <SortIcon /></span>
          </th>
          <th @click="sort(json, 'Особенности')">
            <span>Особенности <SortIcon /></span>
          </th>
          <th @click="sort(json, 'Тип зерна')">
            <span>Тип зерна <SortIcon /></span>
          </th>
          <th @click="sort(json, 'Зернистость')">
            <span>Зернистость <SortIcon /></span>
          </th>

          <th @click="sortDots(json, 'Abraforce')" class="abraforce">
            <span>Abraforce <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P20')" class="abraforce">
            <span>P20 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P24')" class="abraforce">
            <span>P24 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P36')" class="abraforce">
            <span>P36 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P40')" class="abraforce">
            <span>P40 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P50')" class="abraforce">
            <span>P50 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P60')" class="abraforce">
            <span>P60 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P80')" class="abraforce">
            <span>P80 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P100')" class="abraforce">
            <span>P100 <SortIcon /></span>
          </th>

          <th @click="sortDots(json, 'P120')" class="abraforce">
            <span>P120 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P150')" class="abraforce">
            <span>P150 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P180')" class="abraforce">
            <span>P180 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P220')" class="abraforce">
            <span>P220 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P240')" class="abraforce">
            <span>P240 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P280')" class="abraforce">
            <span>P280 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P320')" class="abraforce">
            <span>P320 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P360')" class="abraforce">
            <span>P360 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P400')" class="abraforce">
            <span>P400 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P500')" class="abraforce">
            <span>P500 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P600')" class="abraforce">
            <span>P600 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P800')" class="abraforce">
            <span>P800 <SortIcon /></span>
          </th>

          <th @click="sortDots(json, 'P1000')" class="abraforce">
            <span>P1000 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P1200')" class="abraforce">
            <span>P1200 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P1500')" class="abraforce">
            <span>P1500 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P2000')" class="abraforce">
            <span>P2000 <SortIcon /></span>
          </th>
          <th @click="sortDots(json, 'P2500')" class="abraforce">
            <span>P2500 <SortIcon /></span>
          </th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(row, idx) in json" :key="idx">
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
import SortIcon from "@/components/SortIcon.vue";

const table = ref();
const json = ref([]);

const sheetToJson = async () => {
  const f = await fetch("/src/assets/table.xlsx");
  const ab = await f.arrayBuffer();

  /* parse workbook */
  const workbook = read(ab);

  /* JSON data */
  json.value = utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

  console.log(json.value);

  //raw_data[0]['Бренд'] = 'Новое название'

  /* update data */
  /* html.value = utils.sheet_to_html(utils.json_to_sheet(raw_data)); */

  //console.log(raw_data);
};

const byFieldUp = (fieldName) => {
  return (a, b) => (a[fieldName] > b[fieldName] ? 1 : -1);
};

const byFieldDown = (fieldName) => {
  return (a, b) => (a[fieldName] < b[fieldName] ? 1 : -1);
};

const sort = (arr, field) => {
  if (event.currentTarget.classList.contains("--sort-down")) {
    sortUp(arr, field);
  } else {
    sortDown(arr, field);
  }
};

const sortUp = (arr, field) => {
  arr.sort(byFieldUp(field));

  clearSorting();
  event.currentTarget.classList.add("--sort-up");
};

const sortDown = (arr, field) => {
  arr.sort(byFieldDown(field));

  clearSorting();
  event.currentTarget.classList.add("--sort-down");
};

const sortDots = (arr, field) => {
  if (event.currentTarget.classList.contains("--sort-up")) {
    function comparator(a, b) {
      if (!a[field]) return -1;
      const num_a = +a[field]?.replace(/\s+/g, "");
      const num_b = +b[field]?.replace(/\s+/g, "");
      return num_b - num_a;
    }
    arr.sort(comparator);
    clearSorting();
    event.currentTarget.classList.add("--sort-down");
    console.log("сортировка по убываюнию!");
  } else {
    function comparator(a, b) {
      if (!b[field]) return -1;
      const num_a = +a[field]?.replace(/\s+/g, "");
      const num_b = +b[field]?.replace(/\s+/g, "");
      return num_b - num_a;
    }
    arr.sort(comparator);
    clearSorting();
    event.currentTarget.classList.add("--sort-up");
    console.log("сортировка по возрастанию!");
  }
};

const clearSorting = () => {
  document.querySelectorAll(".sheet-table__header th").forEach((item) => {
    item.classList.remove("--sort-up");
    item.classList.remove("--sort-down");
  });
};

const prohibitToCopy = () => {
  table.value.ondragstart = prohibit;
  table.value.onselectstart = prohibit;
  table.value.oncontextmenu = prohibit;
  function prohibit() {
    return false;
  }
};

onMounted(async () => {
  sheetToJson();
  prohibitToCopy();
});
</script>

<style lang="scss">
.sheet-table {
  $that: &;

  user-select: none;
  -webkit-user-select: none;
  -ms-user-select: none;

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

      span {
        display: flex;
        align-items: center;
        justify-content: center;
      }

      &:hover:not(.--sort-up):not(.--sort-down) {
        .sort-icon {
          opacity: 0.5;
        }
      }

      &.abraforce {
        background-color: rgb(177, 177, 255);
      }
    }
  }
}
</style>
