let GRID = {
    columns: [],
    rows: [],        // full rows (objects)
    viewRows: [],    // filtered rows
    selected: new Set(), // rowId set
    search: "",
    from: "",
    to: ""
};

// C# will call this
function renderMainGrid(payload) {
    GRID.columns = payload.columns || [];
    GRID.rows = payload.rows || [];
    GRID.viewRows = [...GRID.rows];

    // keep selection only for existing rows
    const existing = new Set(GRID.rows.map(r => r.__rowId));
    GRID.selected = new Set([...GRID.selected].filter(id => existing.has(id)));

    applyClientFilters();
}

// expose selection getter for C#
window.__getSelectedRowIds = () => JSON.stringify([...GRID.selected]);

// UI elements
const $ = (id) => document.getElementById(id);

function send(msg) {
    window.chrome?.webview?.postMessage(msg);
}

function sendJson(obj) {
    window.chrome?.webview?.postMessage(JSON.stringify(obj));
}

function parseDateSafe(v) {
    // expects yyyy-mm-dd from inputs
    if (!v) return null;
    const d = new Date(v + "T00:00:00");
    return isNaN(d.getTime()) ? null : d;
}

function rowMatchesSearch(row, text) {
    if (!text) return true;
    const t = text.toLowerCase();
    for (const k of Object.keys(row)) {
        if (k === "__rowId") continue;
        const val = (row[k] ?? "").toString().toLowerCase();
        if (val.includes(t)) return true;
    }
    return false;
}

function rowMatchesDate(row, fromD, toD) {
    if (!fromD && !toD) return true;

    // We assume there is a "Date" column if you want date filtering.
    // If your column name differs, rename it here:
    const dateStr = row["Date"] ?? row["date"] ?? row["DATE"];
    if (!dateStr) return false;

    // Try to parse. If your excel date format is different, tell me and I adjust parser.
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return false;

    if (fromD && d < fromD) return false;
    if (toD) {
        const toEnd = new Date(toD);
        toEnd.setHours(23, 59, 59, 999);
        if (d > toEnd) return false;
    }
    return true;
}

function applyClientFilters() {
    const search = GRID.search.trim();
    const fromD = parseDateSafe(GRID.from);
    const toD = parseDateSafe(GRID.to);

    GRID.viewRows = GRID.rows.filter(r =>
        rowMatchesSearch(r, search) && rowMatchesDate(r, fromD, toD)
    );

    renderTable();
}

function renderTable() {
    const table = $("dataGrid");
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    // Header row
    const trh = document.createElement("tr");

    // select checkbox header
    const thSel = document.createElement("th");
    thSel.textContent = "";
    trh.appendChild(thSel);

    // data headers
    GRID.columns.forEach(c => {
        const th = document.createElement("th");
        th.textContent = c;
        trh.appendChild(th);
    });

    // actions header
    const thA = document.createElement("th");
    thA.textContent = "Actions";
    thA.className = "col-actions";
    trh.appendChild(thA);

    thead.appendChild(trh);

    // Body
    GRID.viewRows.forEach(row => {
        const tr = document.createElement("tr");

        // checkbox
        const tdChk = document.createElement("td");
        const chk = document.createElement("input");
        chk.type = "checkbox";
        chk.checked = GRID.selected.has(row.__rowId);
        chk.addEventListener("change", () => {
            if (chk.checked) GRID.selected.add(row.__rowId);
            else GRID.selected.delete(row.__rowId);
            syncHeaderCheckbox();
        });
        tdChk.appendChild(chk);
        tr.appendChild(tdChk);

        // columns
        // columns
        GRID.columns.forEach(c => {
            const td = document.createElement("td");

            if (c === "Status") {
                const raw = (row[c] ?? "").toString();

                // 🔥 normalize Arabic text (IMPORTANT)
                const val = raw
                    .replace(/\u200F/g, "")   // RTL mark
                    .replace(/\u00A0/g, " ")  // NBSP
                    .replace(/\s+/g, " ")
                    .trim();

                // ✅ ONLY Status decides the color
                const isGreen = val === "غير مدرج";
                const isRed = val === "مدرج";

                const pill = document.createElement("span");
                pill.className = "pill " + (isGreen ? "green" : "red");
                pill.textContent = val;

                pill.style.cursor = "pointer";
                pill.title = "Click to toggle";

                pill.addEventListener("click", () => {
                    const newVal = isGreen ? "مدرج" : "غير مدرج";

                    sendJson({
                        action: "update-status",
                        rowId: row.__rowId,
                        value: newVal
                    });
                });

                td.appendChild(pill);
            }
            else {
                // ✅ other columns are DISPLAY ONLY
                td.textContent = (row[c] ?? "");
            }

            tr.appendChild(td);
        });


        // actions
        const tdA = document.createElement("td");
        tdA.className = "col-actions";

        const btnDel = document.createElement("button");
        btnDel.className = "rowbtn";
        btnDel.textContent = "??";
        btnDel.title = "Delete row";
        btnDel.addEventListener("click", () => {
            if (!confirm("Delete this row?")) return;
            sendJson({ action: "delete-rows", rowIds: [row.__rowId] });
        });

        tdA.appendChild(btnDel);
        tr.appendChild(tdA);

        tbody.appendChild(tr);
    });

    $("rowCount").innerText = `Rows: ${GRID.viewRows.length}`;

    syncHeaderCheckbox();
}

function syncHeaderCheckbox() {
    const chkAll = $("chkAll");
    if (!chkAll) return;

    if (GRID.viewRows.length === 0) {
        chkAll.checked = false;
        chkAll.indeterminate = false;
        return;
    }

    let selectedVisible = 0;
    GRID.viewRows.forEach(r => { if (GRID.selected.has(r.__rowId)) selectedVisible++; });

    chkAll.checked = selectedVisible === GRID.viewRows.length;
    chkAll.indeterminate = selectedVisible > 0 && selectedVisible < GRID.viewRows.length;
}

// -------------
// UI wiring
// -------------
$("btnImport")?.addEventListener("click", () => send("import-excel"));
$("btnExportAll")?.addEventListener("click", () => send("export-all-excel"));
$("btnCompare")?.addEventListener("click", () => send("compare-excel"));

$("btnExportSelectedExcel")?.addEventListener("click", () => {
    if (GRID.selected.size === 0) return alert("No rows selected.");
    send("export-selected-excel");
});

$("btnExportSelectedPdf")?.addEventListener("click", () => {
    if (GRID.selected.size === 0) return alert("No rows selected.");
    send("export-selected-pdf");
});

$("txtSearch")?.addEventListener("input", (e) => {
    GRID.search = e.target.value || "";
    applyClientFilters();
});

$("btnApplyFilters")?.addEventListener("click", () => {
    GRID.from = $("dateFrom")?.value || "";
    GRID.to = $("dateTo")?.value || "";
    applyClientFilters();
});

$("btnClearFilters")?.addEventListener("click", () => {
    GRID.search = "";
    GRID.from = "";
    GRID.to = "";
    if ($("txtSearch")) $("txtSearch").value = "";
    if ($("dateFrom")) $("dateFrom").value = "";
    if ($("dateTo")) $("dateTo").value = "";
    applyClientFilters();
});

$("chkAll")?.addEventListener("change", (e) => {
    const checked = e.target.checked;

    GRID.viewRows.forEach(r => {
        if (checked) GRID.selected.add(r.__rowId);
        else GRID.selected.delete(r.__rowId);
    });

    renderTable();
});

$("btnDeleteSelected")?.addEventListener("click", () => {
    if (GRID.selected.size === 0) return alert("No rows selected.");
    if (!confirm("Delete selected rows?")) return;
    sendJson({ action: "delete-rows", rowIds: [...GRID.selected] });
});
