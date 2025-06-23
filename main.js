// main.js

const APP_KEY = "pivot_app_v1";
let appState = {
  data: [],
  columns: [],
  pivots: [], // {id, title, config}
  activePivotId: null,
  customColumns: [], // {name, mapping}
};

// Data table state for filtering and pagination
let dataTableState = {
  filters: {},
  page: 1,
  pageSize: 50,
};

// State for sorting
let sortState = {
  col: null,
  dir: null, // 'asc' or 'desc'
};

// Add state for hidden columns
let hiddenColumns = JSON.parse(
  localStorage.getItem(APP_KEY + "_hiddenColumns") || "[]"
);

// Helper: get unique values for a column
function getUniqueValues(col) {
  const set = new Set();
  appState.data.forEach((row) => {
    set.add(row[col] || "");
  });
  return Array.from(set).sort();
}

// State for Excel-style filter dropdowns
let filterDropdownState = {
  openCol: null,
  search: "",
  checked: {}, // {col: Set(values)}
  anchorEl: null, // store the icon element
};

function saveState() {
  localStorage.setItem(
    APP_KEY,
    JSON.stringify({
      pivots: appState.pivots,
      activePivotId: appState.activePivotId,
    })
  );
}

function loadState() {
  const saved = localStorage.getItem(APP_KEY);
  if (saved) {
    const parsed = JSON.parse(saved);
    appState.pivots = parsed.pivots || [];
    appState.activePivotId = parsed.activePivotId || null;
  }
}

function initApp() {
  loadState();
  // Add a permanent Data tab if not present
  if (!appState.pivots.some((p) => p.id === "data")) {
    appState.pivots.unshift({
      id: "data",
      title: "Data",
      config: null,
      isData: true,
    });
  }
  if (!appState.activePivotId) appState.activePivotId = "data";
  renderApp();
}

function renderApp() {
  const app = document.getElementById("app");
  app.innerHTML = `
    <div class="app-toolbar">
      <div class="app-toolbar-left">
        <button id="export-pivots-btn">Export Pivots</button>
        <button id="export-pivot-csv-btn">Export Pivot CSV</button>
        <button id="copy-pivot-csv-btn">Copy Pivot TSV</button>

        <label for="import-pivots-input" style="margin-bottom:0;">Import Pivots<input id="import-pivots-input" type="file" accept="application/json" style="display:none"></label>
      </div>
      <div class="app-toolbar-right">
        <button id="add-pivot-btn">+ New Pivot Table</button>
      </div>
    </div>
    <div class="tabs" id="pivot-tabs"></div>
    <div id="main-content"></div>
  `;
  document
    .getElementById("add-pivot-btn")
    .addEventListener("click", handleAddPivot);
  document.getElementById("export-pivots-btn").onclick = exportPivots;
  document.getElementById("export-pivot-csv-btn").onclick = exportPivotCSV;
  document.getElementById("copy-pivot-csv-btn").onclick = copyPivotTSV;
  document.getElementById("import-pivots-input").onchange = function (e) {
    if (e.target.files && e.target.files[0]) {
      importPivots(e.target.files[0]);
    }
  };
  renderTabs();
  renderMainContent();
}

function handleCSVUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: function (results) {
      appState.data = results.data;
      appState.columns = results.meta.fields;
      // Debug output
      const numericCols = (function getNumericColumns() {
        const N = Math.min(1000, appState.data.length);
        const sample = appState.data.slice(0, N);
        return appState.columns.filter((col) => {
          let hasNumeric = false;
          for (let row of sample) {
            const val = row[col];
            if (val !== undefined && val !== null && val !== "") {
              if (!isNaN(Number(val))) {
                hasNumeric = true;
              } else {
                return false;
              }
            }
          }
          return hasNumeric;
        });
      })();
      console.log("Detected numeric columns:", numericCols);
      console.log(
        "First 10 'bs' values:",
        appState.data.slice(0, 10).map((row) => row.bs)
      );
      // Coerce numeric columns to numbers
      (function coerceNumericColumns() {
        appState.data.forEach((row) => {
          numericCols.forEach((col) => {
            if (
              row[col] !== undefined &&
              row[col] !== null &&
              row[col] !== ""
            ) {
              row[col] = Number(row[col]);
            }
          });
        });
      })();
      renderMainContent();
    },
    error: function (err) {
      const mainContent = document.getElementById("main-content");
      if (mainContent) {
        mainContent.innerHTML =
          '<div style="padding:32px;color:#b71c1c;font-size:1.2em;">Error parsing CSV file: ' +
          err.message +
          ". Please ensure it's a valid CSV.</div>";
      }
    },
  });
}

function handleAddPivot() {
  const title = prompt("Enter a title for this pivot table:");
  if (!title) return;
  const id = "pivot_" + Date.now();
  appState.pivots.push({ id, title, config: null });
  appState.activePivotId = id;
  saveState();
  renderTabs();
  renderPivot();
}

function renderTabs() {
  const tabs = document.getElementById("pivot-tabs");
  if (!tabs) return;
  tabs.innerHTML = "";
  appState.pivots.forEach((pivot) => {
    const tab = document.createElement("div");
    tab.className =
      "tab" + (pivot.id === appState.activePivotId ? " active" : "");
    tab.textContent = pivot.title;
    tab.onclick = () => {
      appState.activePivotId = pivot.id;
      saveState();
      renderTabs();
      renderMainContent();
    };
    if (!pivot.isData) {
      // Rename button
      const rename = document.createElement("span");
      rename.className = "rename";
      rename.textContent = "✎";
      rename.style.marginLeft = "6px";
      rename.style.color = "#1976d2";
      rename.style.cursor = "pointer";
      rename.onclick = (e) => {
        e.stopPropagation();
        const newTitle = prompt("Rename pivot table:", pivot.title);
        if (newTitle) {
          pivot.title = newTitle;
          saveState();
          renderTabs();
        }
      };
      tab.appendChild(rename);
      // Close button
      const close = document.createElement("span");
      close.className = "close";
      close.textContent = "×";
      close.onclick = (e) => {
        e.stopPropagation();
        appState.pivots = appState.pivots.filter((p) => p.id !== pivot.id);
        if (appState.activePivotId === pivot.id) {
          appState.activePivotId = "data";
        }
        saveState();
        renderTabs();
        renderMainContent();
      };
      tab.appendChild(close);
    }
    tabs.appendChild(tab);
  });
}

// Remove any existing filter dropdown from body
function removeGlobalFilterDropdown() {
  const existing = document.getElementById("global-excel-filter-dropdown");
  if (existing) existing.remove();
}

function renderDataTable(container) {
  if (!appState.data.length) {
    container.innerHTML = `
      <div style="padding: 40px; text-align: center; background: #f8fafd; border: 1px dashed #d1d5e3; border-radius: 8px;">
        <h2 style="margin-top:0;">Welcome to Pivot Table App</h2>
        <p>To get started, please select your CSV data file.</p>
        <label for="csv-upload-input" style="cursor:pointer; display:inline-block; padding: 10px 20px; background: #1976d2; color: #fff; border-radius: 4px; font-weight: 600;">
          Load CSV File
          <input id="csv-upload-input" type="file" accept=".csv, text/csv" style="display:none">
        </label>
      </div>
    `;
    container.querySelector("#csv-upload-input").onchange = handleCSVUpload;
    removeGlobalFilterDropdown();
    return;
  }
  // --- Show/Hide Columns UI ---
  let showHideBtn = `<button id=\"show-hide-cols-btn\" style=\"margin-bottom:8px;margin-right:12px;\">Show/Hide Columns</button>`;
  let showHideDropdown = `<div id=\"show-hide-cols-dropdown\" style=\"display:none;position:absolute;z-index:1000;background:#fff;border:1px solid #e0e3ea;box-shadow:0 2px 8px rgba(60,80,120,0.08);padding:12px 16px;border-radius:8px;max-height:300px;overflow:auto;\">\n    <b>Show/Hide Columns</b><br>\n    ${appState.columns
    .map(
      (col) =>
        `<label style='display:block;margin:4px 0;'><input type='checkbox' class='shc' value=\"${col}\" ${
          hiddenColumns.includes(col) ? "" : "checked"
        }> ${col}</label>`
    )
    .join(
      ""
    )}\n    <div style='margin-top:8px;text-align:right;'><button id='shc-close-btn'>Close</button></div>\n  </div>`;

  // --- Excel-like Custom Column UI ---
  let customColUI = `<div id="custom-col-ui" style="margin-bottom:16px;padding:12px 16px;background:#f8fafd;border-radius:8px;box-shadow:0 1px 4px rgba(60,80,120,0.04);">
    <b>Add Custom Column (Excel-like Formula)</b><br>
    <div style='margin:6px 0;'>
      <label>New column name: <input id="custom-col-name" type="text" style="margin-left:4px;width:120px;"></label>
    </div>
    <div style='margin:6px 0;'>
      <label>Formula: <input id="custom-col-formula" type="text" style="margin-left:4px;width:340px;" placeholder='e.g. 2025-[dob]'></label>
      <span style='margin-left:12px;'>Insert column: <select id="custom-col-insert-col"><option value="">-- Select --</option>${appState.columns
        .map((col) => `<option value="${col}">${col}</option>`)
        .join("")}</select></span>
    </div>
    <div id="custom-col-error" style="color:#e53935;font-size:11px;margin:4px 0 0 0;"></div>
    <div id="custom-col-preview" style="margin-top:8px;font-size:11px;"></div>
    <button id="custom-col-add-btn" style="margin-top:8px;">Add Custom Column</button>
  </div>`;
  // All Clear Filters button
  let clearAllBtn = `<button id=\"clear-all-filters-btn\" style=\"margin-bottom:8px;background:#e53935;color:#fff;border:none;border-radius:4px;padding:4px 14px;font-size:10px;font-weight:600;box-shadow:0 1px 4px rgba(229,57,53,0.08);cursor:pointer;float:right;\">All Clear Filters</button>`;
  // Filtering logic
  let filtered = appState.data.filter((row) => {
    return appState.columns.every((col) => {
      const checked = filterDropdownState.checked[col];
      if (!checked || checked.size === 0) return true;
      return checked.has(row[col] || "");
    });
  });
  // Sorting logic
  if (sortState.col) {
    filtered = filtered.slice().sort((a, b) => {
      const va = a[sortState.col] || "";
      const vb = b[sortState.col] || "";
      if (!isNaN(va) && !isNaN(vb) && va !== "" && vb !== "") {
        return sortState.dir === "asc" ? va - vb : vb - va;
      }
      return sortState.dir === "asc"
        ? String(va).localeCompare(String(vb))
        : String(vb).localeCompare(String(va));
    });
  }
  // Pagination
  const totalRows = filtered.length;
  const totalPages = Math.ceil(totalRows / dataTableState.pageSize) || 1;
  if (dataTableState.page > totalPages) dataTableState.page = totalPages;
  const startIdx = (dataTableState.page - 1) * dataTableState.pageSize;
  const endIdx = startIdx + dataTableState.pageSize;
  const pageRows = filtered.slice(startIdx, endIdx);
  // Only show visible columns
  const visibleCols = appState.columns.filter(
    (col) => !hiddenColumns.includes(col)
  );
  // Table header with filter and sort icons
  let thead =
    "<tr>" +
    visibleCols
      .map((col) => {
        // SVG icons
        const filterSVG = `<svg class=\"filter-icon\" data-col=\"${col}\" width=\"14\" height=\"14\" viewBox=\"0 0 20 20\" style=\"vertical-align:middle;cursor:pointer;margin-left:4px;\" fill=\"none\" stroke=\"#888\" stroke-width=\"1.5\"><path d=\"M3 5h14M6 9h8M9 13h2\" stroke-linecap=\"round\"/></svg>`;
        const ascActive = sortState.col === col && sortState.dir === "asc";
        const descActive = sortState.col === col && sortState.dir === "desc";
        const sortAscSVG = `<svg class=\"sort-asc\" data-col=\"${col}\" width=\"10\" height=\"10\" viewBox=\"0 0 20 20\" style=\"vertical-align:middle;cursor:pointer;margin-left:4px;\" fill=\"none\" stroke=\"${
          ascActive ? "#1976d2" : "#bbb"
        }\" stroke-width=\"2\"><path d=\"M6 12l4-4 4 4\"/></svg>`;
        const sortDescSVG = `<svg class=\"sort-desc\" data-col=\"${col}\" width=\"10\" height=\"10\" viewBox=\"0 0 20 20\" style=\"vertical-align:middle;cursor:pointer;margin-left:2px;\" fill=\"none\" stroke=\"${
          descActive ? "#1976d2" : "#bbb"
        }\" stroke-width=\"2\"><path d=\"M6 8l4 4 4-4\"/></svg>`;
        return `\n      <th style=\"position:relative;white-space:nowrap;\">\n        <span>${col}</span>\n        ${sortAscSVG}${sortDescSVG}${filterSVG}\n      </th>\n    `;
      })
      .join("") +
    "</tr>";
  // Table body
  let tbody = pageRows
    .map((row) => {
      return (
        "<tr>" +
        visibleCols.map((col) => `<td>${row[col] || ""}</td>`).join("") +
        "</tr>"
      );
    })
    .join("");
  // Pagination controls
  let pagination = `<div style=\"margin:8px 0;display:flex;align-items:center;gap:8px;\">\n    <button id=\"dt-prev\" ${
    dataTableState.page === 1 ? "disabled" : ""
  }>&lt; Prev</button>\n    <span>Page ${
    dataTableState.page
  } / ${totalPages} (${totalRows} rows)</span>\n    <button id=\"dt-next\" ${
    dataTableState.page === totalPages ? "disabled" : ""
  }>Next &gt;</button>\n  </div>`;
  container.innerHTML = `
    <div style='position:relative;'>${showHideBtn}${showHideDropdown}</div>
    ${customColUI}
    ${clearAllBtn}
    <div style=\"overflow-x:auto;clear:both;\">
      <table class=\"data-table\">
        <thead>${thead}</thead>
        <tbody>${tbody}</tbody>
      </table>
    </div>
    ${pagination}
  `;
  // Show/Hide Columns logic
  const shBtn = container.querySelector("#show-hide-cols-btn");
  const shDropdown = container.querySelector("#show-hide-cols-dropdown");
  shBtn.onclick = function (e) {
    e.stopPropagation();
    shDropdown.style.display =
      shDropdown.style.display === "block" ? "none" : "block";
    shDropdown.style.left = shBtn.offsetLeft + "px";
    shDropdown.style.top = shBtn.offsetTop + shBtn.offsetHeight + 2 + "px";
  };
  // Remove global document click event for closing the dropdown
  // Only close dropdown when clicking the 'Close' button
  // Prevent dropdown from closing when clicking inside (including checkboxes and labels)
  shDropdown.addEventListener("mousedown", function (e) {
    e.stopPropagation();
  });
  shDropdown.addEventListener("click", function (e) {
    e.stopPropagation();
  });
  shDropdown.querySelectorAll(".shc").forEach((cb) => {
    cb.onchange = function () {
      const col = cb.value;
      if (cb.checked) {
        hiddenColumns = hiddenColumns.filter((c) => c !== col);
      } else {
        if (!hiddenColumns.includes(col)) hiddenColumns.push(col);
      }
      localStorage.setItem(
        APP_KEY + "_hiddenColumns",
        JSON.stringify(hiddenColumns)
      );
      // Instead of re-rendering the whole data table, just update thead/tbody
      const visibleCols = appState.columns.filter(
        (col) => !hiddenColumns.includes(col)
      );
      // Update thead
      let thead =
        "<tr>" +
        visibleCols
          .map((col) => {
            const filterSVG = `<svg class=\"filter-icon\" data-col=\"${col}\" width=\"14\" height=\"14\" viewBox=\"0 0 20 20\" style=\"vertical-align:middle;cursor:pointer;margin-left:4px;\" fill=\"none\" stroke=\"#888\" stroke-width=\"1.5\"><path d=\"M3 5h14M6 9h8M9 13h2\" stroke-linecap=\"round\"/></svg>`;
            const ascActive = sortState.col === col && sortState.dir === "asc";
            const descActive =
              sortState.col === col && sortState.dir === "desc";
            const sortAscSVG = `<svg class=\"sort-asc\" data-col=\"${col}\" width=\"10\" height=\"10\" viewBox=\"0 0 20 20\" style=\"vertical-align:middle;cursor:pointer;margin-left:4px;\" fill=\"none\" stroke=\"${
              ascActive ? "#1976d2" : "#bbb"
            }\" stroke-width=\"2\"><path d=\"M6 12l4-4 4 4\"/></svg>`;
            const sortDescSVG = `<svg class=\"sort-desc\" data-col=\"${col}\" width=\"10\" height=\"10\" viewBox=\"0 0 20 20\" style=\"vertical-align:middle;cursor:pointer;margin-left:2px;\" fill=\"none\" stroke=\"${
              descActive ? "#1976d2" : "#bbb"
            }\" stroke-width=\"2\"><path d=\"M6 8l4 4 4-4\"/></svg>`;
            return `\n      <th style=\"position:relative;white-space:nowrap;\">\n        <span>${col}</span>\n        ${sortAscSVG}${sortDescSVG}${filterSVG}\n      </th>\n    `;
          })
          .join("") +
        "</tr>";
      // Update tbody
      let filtered = appState.data.filter((row) => {
        return appState.columns.every((col) => {
          const checked = filterDropdownState.checked[col];
          if (!checked || checked.size === 0) return true;
          return checked.has(row[col] || "");
        });
      });
      if (sortState.col) {
        filtered = filtered.slice().sort((a, b) => {
          const va = a[sortState.col] || "";
          const vb = b[sortState.col] || "";
          if (!isNaN(va) && !isNaN(vb) && va !== "" && vb !== "") {
            return sortState.dir === "asc" ? va - vb : vb - va;
          }
          return sortState.dir === "asc"
            ? String(va).localeCompare(String(vb))
            : String(vb).localeCompare(String(va));
        });
      }
      const totalRows = filtered.length;
      const totalPages = Math.ceil(totalRows / dataTableState.pageSize) || 1;
      if (dataTableState.page > totalPages) dataTableState.page = totalPages;
      const startIdx = (dataTableState.page - 1) * dataTableState.pageSize;
      const endIdx = startIdx + dataTableState.pageSize;
      const pageRows = filtered.slice(startIdx, endIdx);
      let tbody = pageRows
        .map((row) => {
          return (
            "<tr>" +
            visibleCols.map((col) => `<td>${row[col] || ""}</td>`).join("") +
            "</tr>"
          );
        })
        .join("");
      const table = container.querySelector(".data-table");
      if (table) {
        table.querySelector("thead").innerHTML = thead;
        table.querySelector("tbody").innerHTML = tbody;
      }
    };
  });
  shDropdown.querySelector("#shc-close-btn").onclick = function () {
    shDropdown.style.display = "none";
  };
  // Filter icon events
  container.querySelectorAll(".filter-icon").forEach((icon) => {
    icon.onclick = (e) => {
      e.stopPropagation();
      const col = icon.getAttribute("data-col");
      if (filterDropdownState.openCol === col) {
        filterDropdownState.openCol = null;
        filterDropdownState.anchorEl = null;
        removeGlobalFilterDropdown();
      } else {
        filterDropdownState.openCol = col;
        filterDropdownState.search = "";
        filterDropdownState.anchorEl = icon;
        setTimeout(() => renderGlobalFilterDropdown(), 0);
      }
    };
  });
  // Sort icon events
  container.querySelectorAll(".sort-asc").forEach((icon) => {
    icon.onclick = (e) => {
      e.stopPropagation();
      const col = icon.getAttribute("data-col");
      if (sortState.col === col && sortState.dir === "asc") {
        sortState.col = null;
        sortState.dir = null;
      } else {
        sortState.col = col;
        sortState.dir = "asc";
      }
      renderDataTable(container);
    };
  });
  container.querySelectorAll(".sort-desc").forEach((icon) => {
    icon.onclick = (e) => {
      e.stopPropagation();
      const col = icon.getAttribute("data-col");
      if (sortState.col === col && sortState.dir === "desc") {
        sortState.col = null;
        sortState.dir = null;
      } else {
        sortState.col = col;
        sortState.dir = "desc";
      }
      renderDataTable(container);
    };
  });
  // Pagination events
  container.querySelector("#dt-prev").onclick = () => {
    if (dataTableState.page > 1) {
      dataTableState.page--;
      renderDataTable(container);
    }
  };
  container.querySelector("#dt-next").onclick = () => {
    if (dataTableState.page < totalPages) {
      dataTableState.page++;
      renderDataTable(container);
    }
  };
  // Close dropdown on outside click
  document.onclick = (e) => {
    if (filterDropdownState.openCol) {
      filterDropdownState.openCol = null;
      filterDropdownState.anchorEl = null;
      removeGlobalFilterDropdown();
      renderDataTable(container);
    }
  };
  // --- Custom Column Add Logic ---
  setTimeout(() => {
    const addBtn = document.getElementById("custom-col-add-btn");
    if (!addBtn) return;
    addBtn.onclick = function () {
      const name = document.getElementById("custom-col-name").value.trim();
      const formula = document
        .getElementById("custom-col-formula")
        .value.trim();
      const insertCol = document.getElementById("custom-col-insert-col").value;
      const errorDiv = document.getElementById("custom-col-error");
      errorDiv.textContent = "";
      if (!name) {
        errorDiv.textContent = "Column name required.";
        return;
      }
      if (!formula) {
        errorDiv.textContent = "Formula required.";
        return;
      }
      if (appState.columns.includes(name)) {
        errorDiv.textContent = "Column name already exists.";
        return;
      }
      // Build a function from the formula
      let fn;
      try {
        // Replace [col] with row[col]
        let expr = formula.replace(
          /\[([^\]]+)\]/g,
          (m, col) => `row[\"${col}\"]`
        );
        // Remove leading = if present
        expr = expr.replace(/^=/, "");
        // eslint-disable-next-line no-new-func
        fn = new Function(
          "row",
          `try { return ${expr}; } catch (e) { return null; }`
        );
      } catch (e) {
        errorDiv.textContent = "Invalid formula.";
        return;
      }
      // Insert column at correct position
      let idx = appState.columns.length;
      if (insertCol && appState.columns.includes(insertCol)) {
        idx = appState.columns.indexOf(insertCol) + 1;
      }
      appState.columns.splice(idx, 0, name);
      // Calculate values for all rows
      appState.data.forEach((row) => {
        let val = fn(row);
        // If result is numeric string, coerce to number
        if (typeof val === "string" && !isNaN(Number(val)) && val !== "")
          val = Number(val);
        // If result is null/undefined, set to empty string
        if (val === null || val === undefined) val = "";
        row[name] = val;
      });
      // Store custom columns with formula and insertCol
      appState.customColumns = appState.customColumns || [];
      appState.customColumns.push({ name, formula, insertCol });
      localStorage.setItem(
        APP_KEY + "_customColumns",
        JSON.stringify(appState.customColumns)
      );
      renderDataTable(container);
      renderPivot();
    };
  }, 0);
}

function renderFilterDropdown(col) {
  const unique = getUniqueValues(col);
  const checked = filterDropdownState.checked[col] || new Set(unique);
  const search = filterDropdownState.search || "";
  const filtered = unique.filter((v) =>
    v.toLowerCase().includes(search.toLowerCase())
  );
  const allChecked = filtered.every((v) => checked.has(v));
  return `
    <div class="excel-filter-dropdown" style="position:absolute;left:0;top:100%;z-index:10;width:220px;background:#fff;border:1px solid #d1d5e3;border-radius:8px;box-shadow:0 4px 16px rgba(60,80,120,0.13);padding:12px 10px 10px 10px;min-width:180px;max-height:320px;overflow:auto;">
      <div style="margin-bottom:8px;font-weight:600;color:#1976d2;">Filter</div>
      <input type="text" class="excel-filter-search" placeholder="Search..." value="${search}" style="width:98%;margin-bottom:8px;padding:3px 6px;font-size:10px;border-radius:4px;border:1px solid #d1d5e3;">
      <div style="max-height:160px;overflow:auto;border:1px solid #f0f0f0;border-radius:4px;background:#fafbfc;">
        <label style="display:block;padding:3px 6px;cursor:pointer;font-size:10px;">
          <input type="checkbox" class="excel-filter-selectall" ${
            allChecked ? "checked" : ""
          } style="margin-right:4px;">(Select All)
        </label>
        ${filtered
          .map(
            (v) => `
          <label style="display:block;padding:3px 6px;cursor:pointer;font-size:10px;">
            <input type="checkbox" class="excel-filter-item" value="${encodeURIComponent(
              v
            )}" ${checked.has(v) ? "checked" : ""} style="margin-right:4px;">${
              v || "<em>(blank)</em>"
            }
          </label>
        `
          )
          .join("")}
      </div>
      <div style="margin-top:8px;text-align:right;">
        <button class="excel-filter-clear" style="font-size:10px;padding:2px 10px;margin-right:4px;">Clear</button>
        <button class="excel-filter-apply" style="font-size:10px;padding:2px 10px;">Apply</button>
      </div>
    </div>
  `;
}

function renderGlobalFilterDropdown() {
  removeGlobalFilterDropdown();
  const col = filterDropdownState.openCol;
  const icon = filterDropdownState.anchorEl;
  if (!col || !icon) return;
  const rect = icon.getBoundingClientRect();
  const drop = document.createElement("div");
  drop.id = "global-excel-filter-dropdown";
  drop.style.position = "absolute";
  drop.style.left = `${rect.left + window.scrollX}px`;
  drop.style.top = `${rect.bottom + window.scrollY}px`;
  drop.style.zIndex = 10000;
  drop.innerHTML = renderFilterDropdown(col);
  document.body.appendChild(drop);
  drop.onclick = (e) => e.stopPropagation();
  // --- Fix Select All logic and checkbox events ---
  // Search box
  const searchInput = drop.querySelector(".excel-filter-search");
  if (searchInput) {
    searchInput.oninput = (ev) => {
      filterDropdownState.search = ev.target.value;
      renderGlobalFilterDropdown();
    };
  }
  // Select All
  const selectAll = drop.querySelector(".excel-filter-selectall");
  if (selectAll) {
    selectAll.onchange = (ev) => {
      const col = filterDropdownState.openCol;
      const unique = getUniqueValues(col);
      const search = filterDropdownState.search || "";
      const filtered = unique.filter((v) =>
        v.toLowerCase().includes(search.toLowerCase())
      );
      if (!filterDropdownState.checked[col])
        filterDropdownState.checked[col] = new Set(unique);
      if (ev.target.checked) {
        filtered.forEach((v) => filterDropdownState.checked[col].add(v));
      } else {
        filtered.forEach((v) => filterDropdownState.checked[col].delete(v));
      }
      renderGlobalFilterDropdown();
    };
  }
  // Item checkboxes
  drop.querySelectorAll(".excel-filter-item").forEach((box) => {
    box.onchange = (ev) => {
      const col = filterDropdownState.openCol;
      const v = decodeURIComponent(ev.target.value);
      if (!filterDropdownState.checked[col])
        filterDropdownState.checked[col] = new Set(getUniqueValues(col));
      if (ev.target.checked) {
        filterDropdownState.checked[col].add(v);
      } else {
        filterDropdownState.checked[col].delete(v);
      }
      renderGlobalFilterDropdown();
    };
  });
  // Clear
  const clearBtn = drop.querySelector(".excel-filter-clear");
  if (clearBtn) {
    clearBtn.onclick = (ev) => {
      const col = filterDropdownState.openCol;
      filterDropdownState.checked[col] = new Set();
      filterDropdownState.openCol = null;
      filterDropdownState.anchorEl = null;
      removeGlobalFilterDropdown();
      renderDataTable(document.getElementById("main-content"));
    };
  }
  // Apply
  const applyBtn = drop.querySelector(".excel-filter-apply");
  if (applyBtn) {
    applyBtn.onclick = (ev) => {
      filterDropdownState.openCol = null;
      filterDropdownState.anchorEl = null;
      removeGlobalFilterDropdown();
      renderDataTable(document.getElementById("main-content"));
    };
  }
  // --- End fix ---
}

// Helper: get string (non-numeric) columns
function getStringColumns() {
  const N = 20;
  const sample = appState.data.slice(0, N);
  return appState.columns.filter((col) => {
    // If all non-empty values in sample are parseable as numbers, treat as numeric
    let isNumeric = true;
    for (let row of sample) {
      const val = row[col];
      if (val !== undefined && val !== null && val !== "") {
        if (isNaN(Number(val))) {
          isNumeric = false;
          break;
        }
      }
    }
    return !isNumeric;
  });
}

// Helper: get numeric columns
function getNumericColumns() {
  const N = Math.min(1000, appState.data.length);
  const sample = appState.data.slice(0, N);
  return appState.columns.filter((col) => {
    let hasNumeric = false;
    for (let row of sample) {
      const val = row[col];
      if (val !== undefined && val !== null && val !== "") {
        if (!isNaN(Number(val))) {
          hasNumeric = true;
        } else {
          return false; // If any value is non-numeric, skip this column
        }
      }
    }
    return hasNumeric;
  });
}

// Helper: is aggregator numeric
function isNumericAggregator(name) {
  const keywords = [
    "sum",
    "avg",
    "mean",
    "median",
    "p25",
    "p50",
    "p75",
    "min",
    "max",
    "stdev",
    "var",
    "fraction",
    "bound",
    "Integer Sum",
  ];
  return keywords.some((k) => name.toLowerCase().includes(k));
}

function renderPivot(container) {
  container.innerHTML = "";
  if (!appState.data.length || !appState.activePivotId) {
    container.innerHTML = "<em>Select or create a pivot table.</em>";
    return;
  }
  const pivot = appState.pivots.find((p) => p.id === appState.activePivotId);
  if (!pivot) {
    container.innerHTML = "<em>Select or create a pivot table.</em>";
    return;
  }
  // Pivot UI
  const pivotDiv = document.createElement("div");
  container.appendChild(pivotDiv);
  // Restore config if available
  const config = pivot.config || {
    rows: [],
    cols: [],
    aggregatorName: "Count",
    vals: [],
    rendererName: "Table",
  };
  // Only allow string columns for rows/cols
  const stringCols = getStringColumns();
  // Only allow numeric columns for vals if aggregator is numeric
  const numericCols = getNumericColumns();
  // Hide non-numeric columns from aggregator value dropdowns if aggregator is numeric
  const hiddenFromAggregators = isNumericAggregator(config.aggregatorName)
    ? appState.columns.filter((col) => !numericCols.includes(col))
    : [];
  // Render PivotTable.js
  $(pivotDiv).pivotUI(
    appState.data,
    {
      rows: config.rows.filter((r) => stringCols.includes(r)),
      cols: config.cols.filter((c) => stringCols.includes(c)),
      aggregatorName: config.aggregatorName,
      vals: config.vals.length
        ? config.vals.filter((v) =>
            isNumericAggregator(config.aggregatorName)
              ? numericCols.includes(v)
              : true
          )
        : isNumericAggregator(config.aggregatorName) && numericCols.length
        ? [numericCols[0]]
        : [],
      rendererName: config.rendererName,
      // Hide numeric fields from the left panel
      hiddenAttributes: appState.columns.filter(
        (col) => !stringCols.includes(col) && !numericCols.includes(col)
      ),
      hiddenFromAggregators,
      // Patch: always show numeric columns in value dropdown for numeric aggregators
      attrDropdown: isNumericAggregator(config.aggregatorName)
        ? numericCols
        : appState.columns,
      onRefresh: function (cfg) {
        // Save only the necessary config
        const saveCfg = {
          rows: cfg.rows,
          cols: cfg.cols,
          aggregatorName: cfg.aggregatorName,
          vals: cfg.vals,
          rendererName: cfg.rendererName,
        };
        pivot.config = saveCfg;
        saveState();
      },
    },
    true
  );
}

// Remove dropdown on tab switch or rerender
function renderMainContent() {
  removeGlobalFilterDropdown();
  const container = document.getElementById("main-content");
  if (appState.activePivotId === "data") {
    renderDataTable(container);
  } else {
    renderPivot(container);
  }
}

document.addEventListener("DOMContentLoaded", initApp);

// Robust custom aggregator registration: wait for PivotTable.js to load
function registerCustomAggregators() {
  if (!window.$ || !$.pivotUtilities) return false;
  var tpl = $.pivotUtilities.aggregatorTemplates;
  var numFmt = $.pivotUtilities.numberFormat();
  tpl.quantileNoZeros = function (q, formatter) {
    return function () {
      return function (data, rowKey, colKey) {
        var vals = [];
        return {
          push: function (record) {
            var v = parseFloat(record);
            if (!isNaN(v) && v !== 0) vals.push(v);
          },
          value: function () {
            if (!vals.length) return null;
            vals.sort(function (a, b) {
              return a - b;
            });
            var pos = (vals.length - 1) * q;
            var base = Math.floor(pos);
            var rest = pos - base;
            if (vals[base + 1] !== undefined) {
              return formatter(
                vals[base] + rest * (vals[base + 1] - vals[base])
              );
            } else {
              return formatter(vals[base]);
            }
          },
          format: formatter,
          numInputs: 0,
        };
      };
    };
  };
  $.pivotUtilities.aggregators = Object.assign(
    {},
    $.pivotUtilities.aggregators,
    {
      "Median (no zeros)": tpl.quantileNoZeros(0.5, numFmt),
      "Q1 (25th percentile, no zeros)": tpl.quantileNoZeros(0.25, numFmt),
      "Q3 (75th percentile, no zeros)": tpl.quantileNoZeros(0.75, numFmt),
      "10th Percentile (no zeros)": tpl.quantileNoZeros(0.1, numFmt),
      "90th Percentile (no zeros)": tpl.quantileNoZeros(0.9, numFmt),
    }
  );
  return true;
}
(function waitForPivotTable() {
  if (!registerCustomAggregators()) {
    setTimeout(waitForPivotTable, 50);
  }
})();

// --- Override PivotTable.js value filter limit to show all values and never show 'too many to list' ---
(function () {
  if (window.$ && $.pivotUtilities && $.pivotUtilities.filter != null) {
    // Remove the default limit (default is 500)
    $.pivotUtilities.filter.maxListSize = 1000000;
    // Patch the UI to never show 'too many to list'
    var origFilter = $.pivotUtilities.filter;
    $.pivotUtilities.filter = function () {
      var filter = origFilter.apply(this, arguments);
      filter.maxListSize = 1000000;
      filter.showTooMany = false;
      return filter;
    };
    // Patch the UI rendering if needed
    if ($.pivotUtilities.filterMenu) {
      var origMenu = $.pivotUtilities.filterMenu;
      $.pivotUtilities.filterMenu = function () {
        var menu = origMenu.apply(this, arguments);
        // Remove any too-many-to-list message
        if (menu && menu.find) {
          menu.find(".pvtFilterBoxMessage").remove();
        }
        return menu;
      };
    }
  }
})();

// --- Force PivotTable.js to always show all filter values, never 'too many to list' ---
(function () {
  if (window.$ && $.pivotUtilities && $.pivotUtilities.filterMenu) {
    $.pivotUtilities.filterMenu = function (opts) {
      var defaults = {
        value: null,
        filter: null,
        sorters: {},
        localeStrings: {},
        values: [],
        maxListSize: 1000000,
      };
      opts = $.extend({}, defaults, opts);
      var values = opts.values || [];
      var filter =
        opts.filter ||
        function () {
          return true;
        };
      var sorters = opts.sorters || {};
      var localeStrings = opts.localeStrings || {};
      var shownValues = values.slice(0, opts.maxListSize);
      var $menu = $("<div>", { class: "pvtFilterBox" });
      $menu.append($("<h4>").text(opts.value + " (" + values.length + ")"));
      // Search box
      var $search = $("<input>", {
        type: "text",
        placeholder: localeStrings.filterSearch || "Search...",
      });
      $menu.append($search);
      // List of checkboxes
      var $list = $("<div>", { class: "pvtFilterList" });
      shownValues.forEach(function (val) {
        var $label = $("<label>");
        var $checkbox = $("<input>", {
          type: "checkbox",
          value: val,
          checked: filter(val),
        });
        $label.append($checkbox).append(document.createTextNode(" " + val));
        $list.append($label);
      });
      $menu.append($list);
      // Buttons
      var $btns = $("<div>", { class: "pvtFilterBoxButtons" });
      $btns.append(
        $("<button>", { class: "pvtButton" }).text(localeStrings.ok || "OK")
      );
      $btns.append(
        $("<button>", { class: "pvtButton" }).text(
          localeStrings.cancel || "Cancel"
        )
      );
      $menu.append($btns);
      return $menu;
    };
  }
})();

// --- Definitive patch: Always show all filter values in PivotTable.js, never 'too many to list' ---
(function () {
  if (window.$ && $.pivotUtilities && $.pivotUtilities.filter != null) {
    $.pivotUtilities.filter = function (opts) {
      var defaults = {
        value: null,
        filter: null,
        sorters: {},
        localeStrings: {},
        values: [],
        maxListSize: 1000000,
      };
      opts = $.extend({}, defaults, opts);
      var values = opts.values || [];
      var filter =
        opts.filter ||
        function () {
          return true;
        };
      var sorters = opts.sorters || {};
      var localeStrings = opts.localeStrings || {};
      var shownValues = values.slice(0, opts.maxListSize);
      var $menu = $("<div>", { class: "pvtFilterBox" });
      $menu.append($("<h4>").text(opts.value + " (" + values.length + ")"));
      // Search box
      var $search = $("<input>", {
        type: "text",
        placeholder: localeStrings.filterSearch || "Search...",
      });
      $menu.append($search);
      // List of checkboxes
      var $list = $("<div>", { class: "pvtFilterList" });
      shownValues.forEach(function (val) {
        var $label = $("<label>");
        var $checkbox = $("<input>", {
          type: "checkbox",
          value: val,
          checked: filter(val),
        });
        $label.append($checkbox).append(document.createTextNode(" " + val));
        $list.append($label);
      });
      $menu.append($list);
      // Buttons
      var $btns = $("<div>", { class: "pvtFilterBoxButtons" });
      $btns.append(
        $("<button>", { class: "pvtButton" }).text(localeStrings.ok || "OK")
      );
      $btns.append(
        $("<button>", { class: "pvtButton" }).text(
          localeStrings.cancel || "Cancel"
        )
      );
      $menu.append($btns);
      return $menu;
    };
  }
})();

// Export pivots to JSON file
function exportPivots() {
  const data = {
    pivots: appState.pivots,
    activePivotId: appState.activePivotId,
  };
  const blob = new Blob([JSON.stringify(data, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "pivots.json";
  a.click();
  URL.revokeObjectURL(url);
}
// Import pivots from JSON file
function importPivots(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = JSON.parse(e.target.result);
      appState.pivots = data.pivots || [];
      appState.activePivotId = data.activePivotId || null;
      saveState();
      renderApp();
    } catch (err) {
      alert("Invalid pivots file");
    }
  };
  reader.readAsText(file);
}

// Add exportPivotCSV and copyPivotCSV functions
function exportPivotCSV() {
  const table = document.querySelector(".pvtTable");
  if (!table) return alert("No pivot table to export.");
  const csv = tableToCSV(table);
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "pivot-table.csv";
  a.click();
  URL.revokeObjectURL(url);
}

function copyPivotTSV() {
  const table = document.querySelector(".pvtTable");
  if (!table) return alert("No pivot table to copy.");
  const tsv = tableToTSV(table);
  navigator.clipboard.writeText(tsv).then(
    () => alert("Pivot table copied as TSV!"),
    () => alert("Failed to copy pivot table.")
  );
}

function tableToCSV(table) {
  let csv = "";
  for (const row of table.rows) {
    const cells = Array.from(row.cells).map((cell) => {
      let text = cell.textContent.replace(/\r?\n|\r/g, " ").replace(/"/g, '""');
      if (text.search(/([",\n])/g) >= 0) text = '"' + text + '"';
      return text;
    });
    csv += cells.join(",") + "\n";
  }
  return csv;
}

function tableToTSV(table) {
  let tsv = "";
  for (const row of table.rows) {
    const cells = Array.from(row.cells).map((cell) => {
      let text = cell.textContent.replace(/\r?\n|\r/g, " ").replace(/\t/g, " ");
      return text;
    });
    tsv += cells.join("\t") + "\n";
  }
  return tsv;
}

// Patch restoreCustomColumns to use formula logic
(function restoreCustomColumns() {
  const customCols = JSON.parse(
    localStorage.getItem(APP_KEY + "_customColumns") || "[]"
  );
  if (customCols.length) {
    customCols.forEach(({ name, formula, insertCol }) => {
      if (!formula) return;
      let fn;
      try {
        let expr = formula.replace(
          /\[([^\]]+)\]/g,
          (m, col) => `row[\"${col}\"]`
        );
        expr = expr.replace(/^=/, "");
        fn = new Function(
          "row",
          `try { return ${expr}; } catch (e) { return null; }`
        );
      } catch (e) {
        return;
      }
      let idx = appState.columns.length;
      if (insertCol && appState.columns.includes(insertCol)) {
        idx = appState.columns.indexOf(insertCol) + 1;
      }
      if (!appState.columns.includes(name))
        appState.columns.splice(idx, 0, name);
      appState.data.forEach((row) => {
        let val = fn(row);
        if (typeof val === "string" && !isNaN(Number(val)) && val !== "")
          val = Number(val);
        if (val === null || val === undefined) val = "";
        row[name] = val;
      });
    });
    appState.customColumns = customCols;
  }
})();
