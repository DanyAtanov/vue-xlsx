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

          <th @click="sort(json, 'Abraforce')" class="abraforce">
            <span>Abraforce <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P20')" class="abraforce">
            <span>P20 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P24')" class="abraforce">
            <span>P24 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P36')" class="abraforce">
            <span>P36 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P40')" class="abraforce">
            <span>P40 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P50')" class="abraforce">
            <span>P50 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P60')" class="abraforce">
            <span>P60 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P80')" class="abraforce">
            <span>P80 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P100')" class="abraforce">
            <span>P100 <SortIcon /></span>
          </th>

          <th @click="sort(json, 'P120')" class="abraforce">
            <span>P120 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P150')" class="abraforce">
            <span>P150 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P180')" class="abraforce">
            <span>P180 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P220')" class="abraforce">
            <span>P220 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P240')" class="abraforce">
            <span>P240 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P280')" class="abraforce">
            <span>P280 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P320')" class="abraforce">
            <span>P320 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P360')" class="abraforce">
            <span>P360 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P400')" class="abraforce">
            <span>P400 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P500')" class="abraforce">
            <span>P500 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P600')" class="abraforce">
            <span>P600 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P800')" class="abraforce">
            <span>P800 <SortIcon /></span>
          </th>

          <th @click="sort(json, 'P1000')" class="abraforce">
            <span>P1000 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P1200')" class="abraforce">
            <span>P1200 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P1500')" class="abraforce">
            <span>P1500 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P2000')" class="abraforce">
            <span>P2000 <SortIcon /></span>
          </th>
          <th @click="sort(json, 'P2500')" class="abraforce">
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
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["Abraforce"] }}
          </td>

          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P20"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P24"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P36"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P40"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P50"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P60"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P80"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P100"] }}
          </td>

          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P120"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P150"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P180"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P220"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P240"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P280"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P320"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P360"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P400"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P500"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P600"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P800"] }}
          </td>

          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P1000"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P1200"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P1500"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P2000"] }}
          </td>
          <td ref="abraforceValue" :data-row="row['Abraforce']">
            {{ row["P2500"] }}
          </td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<script setup>
import { ref, onMounted, watch } from "vue";
import { read, utils } from "xlsx";
import SortIcon from "@/components/SortIcon.vue";

let table = ref();
let json = ref([]);
let commentsList = ref([]);
let abraforceValue = ref(null);
let tableIsLoaded = ref(false);

watch(abraforceValue, () => {
  colorCells();
});

const sheetToJson = async () => {
  const f = await fetch("/src/assets/table.xlsx");
  const ab = await f.arrayBuffer();

  /* parse workbook */
  const workbook = read(ab, {
    cellStyles: true,
  });

  json.value = utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

  const ws = workbook.Sheets[workbook.SheetNames[0]];
  if (!ws) return;
  const ref = utils.decode_range(ws["!ref"]);
  for (let R = 0; R <= ref.e.r; ++R)
    for (let C = 0; C <= ref.e.c; ++C) {
      const addr = utils.encode_cell({ r: R, c: C });
      if (!ws[addr] || !ws[addr].c) continue;
      var comments = ws[addr].c[0].h;
      if (!comments.length) continue;

      commentsList.value.push({ row: ws[addr].v, inProgress: comments });
    }

  console.log(json.value);
  console.log(workbook)

  tableIsLoaded.value = true;
};

const sort = (arr, field) => {
  if (event.currentTarget.classList.contains("--sort-down")) {
    sortUp(arr, field);
  } else {
    sortDown(arr, field);
  }
  setTimeout(() => {
    colorCells();
  }, 0);
};

const sortUp = (arr, field) => {
  arr.sort((a, b) => {
    if (
      a[field]?.toString().toLowerCase() > b[field]?.toString().toLowerCase()
    ) {
      return 1;
    } else if (
      a[field]?.toString().toLowerCase() < b[field]?.toString().toLowerCase()
    ) {
      return -1;
    } else if (
      a[field]?.toString().toLowerCase() == b[field]?.toString().toLowerCase()
    ) {
      return 0;
    } else if (!a[field] && b[field]) {
      return -1;
    } else if (a[field] && !b[field]) {
      return 1;
    }
  });

  clearSorting();
  event.currentTarget.classList.add("--sort-up");
};

const sortDown = (arr, field) => {
  arr.sort((a, b) => {
    if (
      b[field]?.toString().toLowerCase() > a[field]?.toString().toLowerCase()
    ) {
      return 1;
    } else if (
      b[field]?.toString().toLowerCase() < a[field]?.toString().toLowerCase()
    ) {
      return -1;
    } else if (
      a[field]?.toString().toLowerCase() == b[field]?.toString().toLowerCase()
    ) {
      return 0;
    } else if (!a[field] && b[field]) {
      return 1;
    } else if (a[field] && !b[field]) {
      return -1;
    }
  });

  clearSorting();
  event.currentTarget.classList.add("--sort-down");
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

// Костыль для окрашивания ячеек
const colorCells = () => {
  abraforceValue.value.forEach((cell, index) => {
    cell.style.backgroundColor = "white";
  });
  
  abraforceValue.value.forEach((cell, index) => {
    if (cell.innerHTML.trim() !== "") {
      cell.style.backgroundColor = "#D6DEF2";

      for (let i = 0; i < commentsList.value.length; i++) {
        if (cell.dataset.row === commentsList.value[i].row) {
          cell.style.backgroundColor = "#F8E1D1";
          cell.title = "В разработке";
        }
      }
    }
  });
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
      max-width: 200px;
      word-break: break-word;
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
