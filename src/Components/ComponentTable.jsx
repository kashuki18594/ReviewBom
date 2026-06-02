import React, { useEffect, useState, useRef, useMemo } from "react";
import { toast } from "react-toastify";
import * as XLSX from "xlsx";
import { confirmDelete } from "./confirmDelete";
import { confirmRaw } from "./confirmRaw";
import SKUSummaryTable from "./SKUSummaryTable"



const zoho = () => window.ZOHO;
const bomID = () => window.BOMID;

// ─── Paginated fetch ──────────────────────────────────────────────────────────
async function fetchAllPages(entity, query) {
  let page = 1, results = [];
  while (true) {
    const res = await zoho().CRM.API.searchRecord({
      Entity: entity, Type: "criteria", Query: query, page, per_page: 200,
    });
    const records = res?.data ?? [];
    results.push(...records);
    if (records.length < 200) break;
    page++;
  }
  return results;
}

// ─── Excel export helper ──────────────────────────────────────────────────────
// Applies header bold + background colour, auto-width columns
function styleSheet(ws, headerColor = "4472C4") {
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let C = range.s.c; C <= range.e.c; C++) {
    const cellAddr = XLSX.utils.encode_cell({ r: 0, c: C });
    if (!ws[cellAddr]) continue;
    ws[cellAddr].s = {
      font: { bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: headerColor } },
      alignment: { horizontal: "center" },
      border: {
        bottom: { style: "thin", color: { rgb: "CCCCCC" } },
      },
    };
  }
  // Auto column widths
  const colWidths = [];
  for (let C = range.s.c; C <= range.e.c; C++) {
    let max = 10;
    for (let R = range.s.r; R <= range.e.r; R++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell?.v != null) max = Math.max(max, String(cell.v).length + 2);
    }
    colWidths.push({ wch: Math.min(max, 40) });
  }
  ws["!cols"] = colWidths;
}

// ─── Excel export — EY BOM Template Format ───────────────────────────────────
function exportToExcel(rows, skuSummary, dealData) {
  const wb = XLSX.utils.book_new();
  const ws = {};

  // ── Helpers ────────────────────────────────────────────────────────────────
  const cellAddr = (r, c) => XLSX.utils.encode_cell({ r, c }); // 0-indexed

  function setCell(r, c, v, style = {}) {
    const addr = cellAddr(r, c);
    ws[addr] = { v, t: typeof v === "number" ? "n" : "s", s: style };
  }

  const boldStyle = {
    font: { bold: true, name: "Arial", sz: 10 },
  };
  const headerStyle = {
    font: { bold: true, name: "Arial", sz: 10 },
    fill: { fgColor: { rgb: "D9E1F2" } },
    border: {
      bottom: { style: "thin", color: { rgb: "4472C4" } },
    },
  };
  const kitRowStyle = {
    font: { bold: true, name: "Arial", sz: 10 },
    fill: { fgColor: { rgb: "BDD7EE" } },
  };
  const subHeaderStyle = {
    font: { bold: true, name: "Arial", sz: 10 },
    fill: { fgColor: { rgb: "DDEBF7" } },
  };
  const rawMatStyle = {
    font: { name: "Arial", sz: 10 },
  };
  const summaryHeaderStyle = {
    font: { bold: true, name: "Arial", sz: 10 },
    fill: { fgColor: { rgb: "E2EFDA" } },
    border: { bottom: { style: "thin", color: { rgb: "70AD47" } } },
  };
  const summaryTitleStyle = {
    font: { bold: true, name: "Arial", sz: 11 },
  };
  const grandTotalStyle = {
    font: { bold: true, name: "Arial", sz: 10 },
    fill: { fgColor: { rgb: "FFF2CC" } },
  };

  // ── Row cursor ─────────────────────────────────────────────────────────────
  let R = 0; // current row (0-indexed)

  // ── Section 1 · Header block ───────────────────────────────────────────────
  // Row 0 — blank
  R = 1; // start at row index 1 (template row 2)

  const projectName = dealData?.Deal_Name || dealData?.Account_Name?.name || "";
  const architect = dealData?.Owner?.name || "";
  const glassSpecs = dealData?.Glass_Specs || "";
  const finish = dealData?.Finish || "";

  const labelStyle = { font: { bold: true, name: "Arial", sz: 10 } };


  R += 5; // blank row after header block

  // ── Section 2 · Table column headers ─────────────────────────────────────
  setCell(R, 0, "#", headerStyle);
  setCell(R, 1, "SKU", headerStyle);
  setCell(R, 2, "Name of Component", headerStyle);
  setCell(R, 3, "QTY", headerStyle);
  R++;

  // ── Section 3 · BOM blocks ────────────────────────────────────────────────
  rows.forEach((bom, bomIdx) => {
    // Kit header row
    setCell(R, 0, bomIdx + 1, kitRowStyle);
    setCell(R, 1, bom.SKU || "", kitRowStyle);
    setCell(R, 2, bom.Kit_Name || bom.Po_Name || "", kitRowStyle);
    setCell(R, 3, bom.Qty ?? "", kitRowStyle);
    R++;

    // // Raw-material sub-header
    // // setCell(R, 0, "", subHeaderStyle);
    // // setCell(R, 1, "SKU", subHeaderStyle);
    // // setCell(R, 2, "Name of Component / Description", subHeaderStyle);
    // // setCell(R, 3, "Round Off", subHeaderStyle);
    // R++;

    // Blank spacer (matches template)
    // R++;

    // Raw material rows
    bom.rawMaterials.forEach((mat) => {
      setCell(R, 0, "", rawMatStyle);
      setCell(R, 1, mat.SKU || "", rawMatStyle);
      setCell(R, 2, mat.Description || "", rawMatStyle);
      setCell(R, 3, mat.Round_Off ?? "", { ...rawMatStyle, t: "n" });
      R++;
    });

    // Blank row between BOM blocks
    R++;
  });

  // ── Section 4 · SKU Summary table ────────────────────────────────────────
  R += 1; // extra spacing
  setCell(R, 1, "SKU Merged Summary");
  R++;

  // Summary header
  setCell(R, 1, "SKU", summaryHeaderStyle);
  setCell(R, 2, "Name of Component", summaryHeaderStyle);
  setCell(R, 3, "Round Off", summaryHeaderStyle);
  R++;

  // Summary data rows
  skuSummary.forEach((item) => {
    setCell(R, 1, item.SKU || "", {});
    setCell(R, 2, item.Description || "", {});
    setCell(R, 3, item.totalRoundOff ?? "", { t: "n" });
    R++;
  });

  // Grand total row
  const grandTotal = skuSummary.reduce((s, r) => s + (r.totalRoundOff || 0), 0);
  setCell(R, 1, "Grand Total", grandTotalStyle);
  setCell(R, 2, "", grandTotalStyle);
  setCell(R, 3, grandTotal, { ...grandTotalStyle, t: "n" });
  R++;

  // ── Set sheet range & column widths ────────────────────────────────────────
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: R, c: 5 } });
  ws["!cols"] = [
    { wch: 4 },   // A — #
    { wch: 20 },   // B — SKU
    { wch: 55 },   // C — Name of Component
    { wch: 12 },   // D — QTY / Round Off / Total
    { wch: 14 },   // E — floor / Final Qty
    { wch: 8 },   // F — UOM
  ];

  XLSX.utils.book_append_sheet(wb, ws, "BOM");

  const date = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `BOM_Export_${date}.xlsx`, { bookType: "xlsx", cellStyles: true });
  toast.success("Excel exported successfully!");
}

// ─── EditableField ────────────────────────────────────────────────────────────
function EditableField({ isEditing, value, field, type = "text", onChange, editable = true }) {
  if (!isEditing || !editable) return <span>{value ?? "—"}</span>;
  return (
    <input
      type={type}
      value={value ?? ""}
      onChange={(e) => onChange(field, e.target.value)}
      className="border rounded px-2 py-1 w-full"
    />
  );
}

// ─── RawMaterialRow ───────────────────────────────────────────────────────────
function RawMaterialRow({ material, isEditing, onEdit, onSave, onCancel, onDelete, onChange, saving }) {
  const ef = (field, type = "text", editable = true) => (
    <EditableField isEditing={isEditing} value={material[field]} field={field} type={type} editable={editable} onChange={onChange} />
  );
  return (
    <tr className="bg-white">
      <td className="border p-2 pl-12">{ef("SKU")}</td>
      <td className="border p-2">{ef("Description", "text", false)}</td>
      <td className="border p-2">{ef("Po_Name", "text", false)}</td>
      <td className="border p-2">{ef("Round_Off", "number")}</td>
      <td className="border p-2">
        <div className="flex gap-2 flex-wrap">
          {!isEditing && <button onClick={onEdit} className="bg-blue-500 text-white px-2 py-1 rounded text-sm">Edit</button>}
          {isEditing && <>
            <button onClick={onSave} disabled={saving} className="bg-green-500 text-white px-2 py-1 rounded text-sm disabled:opacity-50">{saving ? "Saving…" : "Save"}</button>
            <button onClick={onCancel} disabled={saving} className="bg-gray-400 text-white px-2 py-1 rounded text-sm">Cancel</button>
          </>}
          <button onClick={() => confirmDelete(onDelete)} className="bg-red-500 text-white px-2 py-1 rounded text-sm">Delete</button>
        </div>
      </td>
    </tr>
  );
}

// ─── BOMRow ───────────────────────────────────────────────────────────────────
function BOMRow({ bom, isEditing, expanded, onToggle, onEdit, onSave, onCancel, onDelete, onAddRaw, onChange, saving }) {
  const ef = (field, type = "text") => (
    <EditableField isEditing={isEditing} value={bom[field]} field={field} type={type} onChange={onChange} />
  );
  return (
    <tr className="bg-gray-100">
      <td className="border p-2">
        <div className="flex items-center gap-2">
          <button onClick={onToggle} className="w-6 h-6 bg-gray-300 rounded flex items-center justify-center text-xs">
            {expanded ? "▼" : "▶"}
          </button>
          {ef("SKU")}
        </div>
        <div className="text-xs text-gray-500 mt-1 ml-8">{bom.rawMaterials.length} Raw Material(s)</div>
      </td>
      <td className="border p-2">{ef("Kit_Name")}</td>
      <td className="border p-2">{ef("Po_Name")}</td>
      <td className="border p-2">{ef("Qty", "number")}</td>
      <td className="border p-2">
        <div className="flex gap-2 flex-wrap">
          <button onClick={() => confirmRaw(onAddRaw)} className="bg-purple-500 text-white px-2 py-1 rounded text-sm">+ Add</button>
          {!isEditing && <button onClick={onEdit} className="bg-blue-500 text-white px-2 py-1 rounded text-sm">Edit</button>}
          {isEditing && <>
            <button onClick={onSave} disabled={saving} className="bg-green-500 text-white px-2 py-1 rounded text-sm disabled:opacity-50">{saving ? "Saving…" : "Save"}</button>
            <button onClick={onCancel} disabled={saving} className="bg-gray-400 text-white px-2 py-1 rounded text-sm">Cancel</button>
          </>}
          <button onClick={() => confirmDelete(onDelete)} className="bg-red-500 text-white px-2 py-1 rounded text-sm">Delete</button>
        </div>
      </td>
    </tr>
  );
}


// ─── Main Component ───────────────────────────────────────────────────────────
export default function ComponentTable({ users }) {
  const [rows, setRows] = useState([]);
  const [dealData, setDealData] = useState(null);
  const [finalBom, setFinalBom] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [expanded, setExpanded] = useState({});
  const [editTarget, setEditTarget] = useState(null);
  const backup = useRef(null);

  // ── SKU Summary (live derived) ────────────────────────────────────────────
  const skuSummary = useMemo(() => {
    const allRaw = rows.flatMap((bom) => bom.rawMaterials);
    const map = new Map();
    for (const mat of allRaw) {
      const sku = (mat.SKU ?? "").trim();
      if (!sku) continue;
      const roundOff = parseFloat(mat.Round_Off) || 0;
      if (map.has(sku)) {
        map.get(sku).totalRoundOff += roundOff;
      } else {
        map.set(sku, { SKU: sku, Description: mat.Description || mat.Name || "", totalRoundOff: roundOff });
      }
    }
    return Array.from(map.values()).sort((a, b) => a.SKU.localeCompare(b.SKU));
  }, [rows]);

  // ── Initial load (once) ───────────────────────────────────────────────────
  useEffect(() => {
    const t = setTimeout(initialLoad, 1000);
    return () => clearTimeout(t);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  async function initialLoad() {
    try {
      setLoading(true);
      const dealRes = await zoho().CRM.API.getRecord({ Entity: "Deals", RecordID: bomID() });
      const deal = dealRes?.data?.[0];
      if (deal) setDealData(deal);

      const bomList = (
        await fetchAllPages("BOM", `((Deals:equals:${bomID()}) and (BOM_Stage:equals:Initial BOM))`)
      ).filter((r) => r.BOM_Status !== "Cancelled");

      const finalBomList = (
        await fetchAllPages("BOM", `((Deals:equals:${bomID()}) and (BOM_Stage:equals:Final BOM))`)
      ).filter((r) => r.BOM_Status !== "Cancelled");

      setFinalBom(finalBomList);



      const rawPerBOM = await Promise.all(
        bomList.map((bom) => fetchAllPages("RawMaterials", `(BOM:equals:${bom.id})`))
      );

      setRows(bomList.map((bom, i) => ({
        ...bom,
        rawMaterials: rawPerBOM[i].map((r) => ({
          ...r,
          _bomId: bom.id,
          Description: r.Description ?? r.Name ?? "",
          Po_Name: r.Po_Name ?? "",
          Round_Off: r.Round_Off ?? 1,
          SKU: r.SKU ?? "",
        })),
      })));
    } catch (err) {
      console.error(err);
      toast.error(err?.message || "Failed to load data");
    } finally {
      setLoading(false);
    }
  }

  // ── Edit helpers ──────────────────────────────────────────────────────────
  function startEdit(bomIndex, rawIndex = null) {
    backup.current = JSON.parse(JSON.stringify(rows));
    setEditTarget({ bomIndex, rawIndex });
  }
  function cancelEdit() {
    if (backup.current) setRows(backup.current);
    backup.current = null;
    setEditTarget(null);
  }
  function handleFieldChange(bomIndex, rawIndex, field, value) {
    setRows((prev) =>
      prev.map((bom, bi) => {
        if (bi !== bomIndex) return bom;
        if (rawIndex === null) return { ...bom, [field]: value };
        return { ...bom, rawMaterials: bom.rawMaterials.map((mat, ri) => ri !== rawIndex ? mat : { ...mat, [field]: value }) };
      })
    );
  }

  // ── Save ──────────────────────────────────────────────────────────────────
  async function handleSave() {
    const { bomIndex, rawIndex } = editTarget;
    setSaving(true);
    try {
      if (rawIndex === null) await saveBOM(bomIndex);
      else await saveRawMaterial(bomIndex, rawIndex);
      backup.current = null;
      setEditTarget(null);
    } catch (err) {
      console.error(err);
      toast.error(err?.message || "Save failed");
    } finally {
      setSaving(false);
    }
  }

  async function saveBOM(bomIndex) {
    const row = rows[bomIndex];
    const payload = { SKU: row.SKU, Kit_Name: row.Kit_Name, Po_Name: row.Po_Name, Qty: row.Qty };
    if (row._isNew) {
      const res = await zoho().CRM.API.insertRecord({
        Entity: "BOM",
        APIData: { ...payload, Deals: { id: bomID() }, Owner: { id: users?.users?.[0]?.id }, Account_Name: { id: dealData?.Account_Name?.id }, BOM_Stage: "Initial BOM" },
        Trigger: ["workflow"],
      });
      const createdId = res?.data?.[0]?.details?.id;
      if (!createdId) throw new Error("BOM create failed — no ID returned from Zoho");
      setRows((prev) => prev.map((bom, bi) => bi !== bomIndex ? bom : { ...bom, id: createdId, _isNew: false }));
      toast.success("BOM Created Successfully");
    } else {
      await zoho().CRM.API.updateRecord({ Entity: "BOM", APIData: { id: row.id, ...payload }, Trigger: ["workflow"] });
      toast.success("BOM Updated Successfully");
    }
  }

  async function saveRawMaterial(bomIndex, rawIndex) {
    const row = rows[bomIndex].rawMaterials[rawIndex];
    const parentId = rows[bomIndex].id;

    // ✅ Duplicate SKU check — block if same SKU already exists in this BOM at a different index
    const isDuplicate = rows[bomIndex].rawMaterials.some(
      (mat, ri) => ri !== rawIndex && mat.SKU?.trim().toLowerCase() === row.SKU?.trim().toLowerCase()
    );

    if (isDuplicate) {
      toast.error(
        `"${row.SKU}" is already added in this BOM. Please update the quantity instead of adding a duplicate.`
      );
      // ✅ Remove the unsaved duplicate row from UI if it's a new row
      if (row._isNew) {
        setRows((prev) =>
          prev.map((bom, bi) =>
            bi !== bomIndex
              ? bom
              : {
                ...bom,
                rawMaterials: bom.rawMaterials.filter((_, ri) => ri !== rawIndex),
              }
          )
        );
      }
      setEditTarget(null);
      return;
    }

    const masterRes = await zoho().CRM.API.searchRecord({ Entity: "Products", Type: "criteria", Query: `(Product_Code:equals:${row.SKU})` });
    const product = masterRes?.data?.[0];
    if (!product) { toast.error("SKU not found in Product Master"); return; }

    const description = product?.Description || product?.Product_Description || row.Description;

    // Immediate UI update with product description
    setRows((prev) =>
      prev.map((bom, bi) =>
        bi !== bomIndex ? bom : {
          ...bom,
          rawMaterials: bom.rawMaterials.map((mat, ri) =>
            ri !== rawIndex ? mat : { ...mat, Description: description }
          ),
        }
      )
    );

    const payload = { SKU: row.SKU, Description: description, Po_Name: row.Po_Name, Round_Off: row.Round_Off };

    if (row._isNew) {
      const res = await zoho().CRM.API.insertRecord({
        Entity: "RawMaterials",
        APIData: { ...payload, BOM: { id: parentId }, Name: `${rawIndex + 1}`, S_No: `${rawIndex + 1}` },
        Trigger: ["workflow"],
      });
      const createdId = res?.data?.[0]?.details?.id;
      if (!createdId) throw new Error("Raw Material create failed — no ID returned from Zoho");
      setRows((prev) =>
        prev.map((bom, bi) =>
          bi !== bomIndex ? bom : {
            ...bom,
            rawMaterials: bom.rawMaterials.map((mat, ri) =>
              ri !== rawIndex ? mat : { ...mat, id: createdId, _isNew: false }
            ),
          }
        )
      );
      toast.success("Raw Material Created Successfully");
    } else {
      await zoho().CRM.API.updateRecord({ Entity: "RawMaterials", APIData: { id: row.id, ...payload }, Trigger: ["workflow"] });
      toast.success("Raw Material Updated Successfully");
    }
  }

  // ── Add ───────────────────────────────────────────────────────────────────
  function addBOM() {
    const newBOM = { id: `_new_bom_${Date.now()}`, _isNew: true, SKU: "", Kit_Name: "", Po_Name: "", Qty: 1, rawMaterials: [] };
    setRows((prev) => {
      const next = [...prev, newBOM];
      backup.current = JSON.parse(JSON.stringify(prev));
      setEditTarget({ bomIndex: next.length - 1, rawIndex: null });
      return next;
    });
  }

  function addRawMaterial(bomIndex) {
    const newMat = { id: `_new_raw_${Date.now()}`, _isNew: true, SKU: "", Description: "", Po_Name: "", Round_Off: 1 };
    setRows((prev) => {
      const next = prev.map((bom, bi) =>
        bi !== bomIndex ? bom : { ...bom, rawMaterials: [...bom.rawMaterials, newMat] }
      );
      const newRawIndex = next[bomIndex].rawMaterials.length - 1;
      setExpanded((e) => ({ ...e, [next[bomIndex].id]: true }));
      backup.current = JSON.parse(JSON.stringify(prev));
      setEditTarget({ bomIndex, rawIndex: newRawIndex });
      return next;
    });
  }

  // ── Delete ────────────────────────────────────────────────────────────────
  async function removeBOM(bomIndex) {
    const bom = rows[bomIndex];
    try {
      if (!bom._isNew) await zoho().CRM.API.deleteRecord({ Entity: "BOM", RecordID: bom.id });
      setRows((prev) => prev.filter((_, bi) => bi !== bomIndex));
      toast.success("BOM deleted");
    } catch (err) { toast.error(err?.message || "Failed to delete BOM"); }
  }

  async function removeRawMaterial(bomIndex, rawIndex) {
    const mat = rows[bomIndex].rawMaterials[rawIndex];
    try {
      if (!mat._isNew) await zoho().CRM.API.deleteRecord({ Entity: "RawMaterials", RecordID: mat.id });
      setRows((prev) =>
        prev.map((bom, bi) =>
          bi !== bomIndex ? bom : { ...bom, rawMaterials: bom.rawMaterials.filter((_, ri) => ri !== rawIndex) }
        )
      );
      toast.success("Raw Material deleted");
    } catch (err) { toast.error(err?.message || "Failed to delete Raw Material"); }
  }

  console.log({ finalBom })

  // ── Render ────────────────────────────────────────────────────────────────
  if (loading) {
    return (
      <div className="h-full w-full flex justify-center items-center py-20">
        <div className="h-8 w-8 animate-spin rounded-full border-4 border-blue-600 border-t-transparent" />
      </div>
    );
  }

  return (
    <div className="p-5">

      {/* ── Toolbar ── */}
      <div className="mb-4 flex items-center gap-3 flex-wrap">
        <button onClick={addBOM} className="bg-blue-600 text-white px-4 py-2 rounded hover:bg-blue-700 text-sm">
          + Add BOM
        </button>

        {/* Export button — disabled when nothing to export */}
        <button
          onClick={() => exportToExcel(rows, skuSummary)}
          disabled={rows.length === 0}
          className="flex items-center gap-2 bg-green-600 text-white px-4 py-2 rounded hover:bg-green-700 text-sm disabled:opacity-40 disabled:cursor-not-allowed"
        >
          {/* spreadsheet icon */}
          <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <polyline points="14 2 14 8 20 8" />
            <line x1="8" y1="13" x2="16" y2="13" />
            <line x1="8" y1="17" x2="16" y2="17" />
            <polyline points="10 9 9 9 8 9" />
          </svg>
          Export Excel
        </button>
      </div>

      {/* ── Main Table ── */}
      <table className="w-full border-collapse border">
        <thead className="bg-gray-200">
          <tr>
            <th className="border p-2 text-left">SKU</th>
            <th className="border p-2 text-left">Name of Component / Description</th>
            <th className="border p-2 text-left">Po Name</th>
            <th className="border p-2 text-left">Qty / Round Off</th>
            <th className="border p-2 text-left">Actions</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((bom, bomIndex) => {
            const isBOMEditing = editTarget?.bomIndex === bomIndex && editTarget?.rawIndex === null;
            return (
              <React.Fragment key={bom.id}>
                <BOMRow
                  bom={bom}
                  isEditing={isBOMEditing}
                  expanded={!!expanded[bom.id]}
                  saving={saving}
                  onToggle={() => setExpanded((e) => ({ ...e, [bom.id]: !e[bom.id] }))}
                  onEdit={() => startEdit(bomIndex)}
                  onSave={handleSave}
                  onCancel={cancelEdit}
                  onDelete={() => removeBOM(bomIndex)}
                  onAddRaw={() => addRawMaterial(bomIndex)}
                  onChange={(field, value) => handleFieldChange(bomIndex, null, field, value)}
                />
                {expanded[bom.id] &&
                  bom.rawMaterials.map((mat, rawIndex) => {
                    const isRawEditing = editTarget?.bomIndex === bomIndex && editTarget?.rawIndex === rawIndex;
                    return (
                      <RawMaterialRow
                        key={mat.id}
                        material={mat}
                        isEditing={isRawEditing}
                        saving={saving}
                        onEdit={() => startEdit(bomIndex, rawIndex)}
                        onSave={handleSave}
                        onCancel={cancelEdit}
                        onDelete={() => removeRawMaterial(bomIndex, rawIndex)}
                        onChange={(field, value) => handleFieldChange(bomIndex, rawIndex, field, value)}
                      />
                    );
                  })}
              </React.Fragment>
            );
          })}
          {rows.length === 0 && (
            <tr>
              <td colSpan={5} className="border p-4 text-center text-gray-400">
                No BOMs found. Click "+ Add BOM" to create one.
              </td>
            </tr>
          )}
        </tbody>
      </table>

      {/* ── SKU Summary ── */}
      <SKUSummaryTable
        summary={skuSummary}
        dealData={dealData}
        dealId={bomID()}
        existingBOMId={finalBom ? finalBom[0]?.id : null}
        users={users} />
    </div>
  );
}