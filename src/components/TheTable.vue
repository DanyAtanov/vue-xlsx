<template>
  <div class="sheet-table">
    <div class="sheet-table__tabs">
      <div
        @click="tabClick(index)"
        v-for="(tab, index) in tableNames"
        :key="index"
        ref="tab"
        class="sheet-table__tab-title"
      >
        {{ tab }}
      </div>
    </div>
    <div
      v-for="(json, i) in jsonList"
      :key="i"
      ref="table"
      class="sheet-table__table"
    >
      <h2>{{ tableNames[i] }}</h2>
      <table>
        <thead class="sheet-table__table-header">
          <tr>
            <th>1</th>
            <th
              v-for="(item, index) in headers[i]"
              :key="index"
              @click="sort(json, item)"
              :class="item.startsWith('P') ? 'abraforce' : ''"
            >
              <span>{{ item }} <SortIcon /></span>
            </th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(row, idx) in json" :key="idx">
            <td>{{ idx + 2 }}</td>
            <td
              v-for="(item, index) in headers[i]"
              :key="index"
              :ref="item.startsWith('P') ? 'abraforceValue' : ''"
              :data-row="item.startsWith('P') ? row['Abraforce'] : ''"
            >
              {{ json[idx][item] ? json[idx][item] : "" }}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<script setup>
import { ref, onMounted, watch } from "vue";
import { read, utils } from "xlsx";
import SortIcon from "@/components/SortIcon.vue";

let table = ref();
let jsonList = ref([]);
let tableList = ref([]);
let tableNames = ref([]);
let commentsList = ref([]);
let headers = ref([]);
let tab = ref([]);
let abraforceValue = ref(null);
let tableIsLoaded = ref(false);

const sheetToJson = async () => {
  const f = await fetch("/src/assets/table.xlsx");
  const ab = await f.arrayBuffer();

  /* parse workbook */
  const workbook = read(ab, {
    cellStyles: true,
  });

  tableList.value = Object.entries(workbook.Sheets).map((e) => ({
    [e[0]]: e[1],
  }));
  tableList.value = [...tableList.value];

  tableList.value.forEach((table, index) => {
    jsonList.value.push(utils.sheet_to_json(table[workbook.SheetNames[index]]));

    collectTitles(jsonList.value[index]);
  });

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

  getTableName(tableList.value);
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
  document.querySelectorAll(".sheet-table__table-header th").forEach((item) => {
    item.classList.remove("--sort-up");
    item.classList.remove("--sort-down");
  });
};

/* const prohibitToCopy = () => {
  table.value.ondragstart = prohibit;
  table.value.onselectstart = prohibit;
  table.value.oncontextmenu = prohibit;
  function prohibit() {
    return false;
  }
}; */

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

const collectTitles = (arr) => {
  let uniqueKeys = new Set();

  arr.forEach((obj) => {
    const keys = Object.keys(obj);
    keys.forEach((key) => uniqueKeys.add(key));
  });

  headers.value.push(uniqueKeys);
};

const getTableName = (arr) => {
  let titles = new Set();

  arr.forEach((obj) => {
    const keys = Object.keys(obj);
    keys.forEach((key) => titles.add(key));
  });

  tableNames.value = [...titles];
};

const tabClick = (index) => {
  console.log(index);
  console.log(table.value[index]);

  table.value.forEach((item, index) => {
    item.classList.remove("--is-active");
    tab.value[index].classList.remove("--is-active");
  });

  table.value[index].classList.add("--is-active");
  tab.value[index].classList.add("--is-active");
};

onMounted(async () => {
  sheetToJson();
  //prohibitToCopy();

  console.log(commentsList.value);
});

watch(abraforceValue, () => {
  colorCells();
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

  &__tabs {
    display: flex;
  }

  &__tab-title {
    cursor: pointer;
    font-weight: bold;

    & + & {
      margin-left: 20px;
    }

    &.--is-active {
      color: #e54934;
    }

    &:hover {
      text-decoration: underline;
    }
  }

  &__table {
    display: none;

    &.--is-active {
      display: block;
    }
  }

  &__table-header {
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
