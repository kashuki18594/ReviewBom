import React, { useEffect, useState, useRef, useMemo } from "react";
import { toast } from "react-toastify";
import * as XLSX from "xlsx";
import { confirmDelete } from "./confirmDelete";
import { confirmRaw } from "./confirmRaw";
import SKUSummaryTable from "./SKUSummaryTable";
import { confirmRejectBom } from "./confirmRejectBom";
import { reviewConfirm } from "./reviewConfirm";

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

// ─── Excel export ─────────────────────────────────────────────────────────────
function exportToExcel(rows, skuSummary, dealData) {
  const wb = XLSX.utils.book_new();
  const ws = {};

  const cellAddr = (r, c) => XLSX.utils.encode_cell({ r, c });

  function setCell(r, c, v, style = {}) {
    const addr = cellAddr(r, c);
    ws[addr] = { v, t: typeof v === "number" ? "n" : "s", s: style };
  }

  const headerStyle = {
    font: { bold: true, name: "Arial", sz: 10 },
    fill: { fgColor: { rgb: "D9E1F2" } },
    border: { bottom: { style: "thin", color: { rgb: "4472C4" } } },
  };
  const kitRowStyle = {
    font: { bold: true, name: "Arial", sz: 10 },
    fill: { fgColor: { rgb: "BDD7EE" } },
  };
  const rawMatStyle = { font: { name: "Arial", sz: 10 } };
  const summaryHeaderStyle = {
    font: { bold: true, name: "Arial", sz: 10 },
    fill: { fgColor: { rgb: "E2EFDA" } },
    border: { bottom: { style: "thin", color: { rgb: "70AD47" } } },
  };
  const grandTotalStyle = {
    font: { bold: true, name: "Arial", sz: 10 },
    fill: { fgColor: { rgb: "FFF2CC" } },
  };

  let R = 0;

  // Table headers
  setCell(R, 0, "#", headerStyle);
  setCell(R, 1, "SKU", headerStyle);
  setCell(R, 2, "Name of Component", headerStyle);
  setCell(R, 3, "Single QTY", headerStyle);
  setCell(R, 4, "Order QTY", headerStyle);
  R++;

  rows.forEach((bom, bomIdx) => {
    setCell(R, 0, bomIdx + 1, kitRowStyle);
    setCell(R, 1, bom.SKU || "", kitRowStyle);
    setCell(R, 2, bom.Kit_Name || bom.Po_Name || "", kitRowStyle);
    setCell(R, 3, "", kitRowStyle);
    setCell(R, 4, parseFloat(bom.Qty) || "", kitRowStyle);
    R++;

    bom.rawMaterials.forEach((mat) => {
      const singleQty = parseFloat(mat.Single_Quantity) || 0;
      const orderQty = parseFloat(mat.Round_Off) || 0;
      setCell(R, 0, "", rawMatStyle);
      setCell(R, 1, mat.SKU || "", rawMatStyle);
      setCell(R, 2, mat.Description || "", rawMatStyle);
      setCell(R, 3, singleQty, { ...rawMatStyle, t: "n" });
      setCell(R, 4, orderQty, { ...rawMatStyle, t: "n" });
      R++;
    });

    R++;
  });

  // SKU Summary
  R += 1;
  setCell(R, 1, "SKU Merged Summary", { font: { bold: true, sz: 11 } });
  R++;
  setCell(R, 1, "SKU", summaryHeaderStyle);
  setCell(R, 2, "Name of Component", summaryHeaderStyle);
  setCell(R, 3, "Single QTY", summaryHeaderStyle);
  setCell(R, 4, "Order QTY", summaryHeaderStyle);
  R++;

  skuSummary.forEach((item) => {
    setCell(R, 1, item.SKU || "", {});
    setCell(R, 2, item.Description || "", {});
    setCell(R, 3, item.totalSingleQty ?? "", { t: "n" });
    setCell(R, 4, item.totalOrderQty ?? "", { t: "n" });
    R++;
  });

  const grandTotalSingle = skuSummary.reduce((s, r) => s + (r.totalSingleQty || 0), 0);
  const grandTotalOrder = skuSummary.reduce((s, r) => s + (r.totalOrderQty || 0), 0);
  setCell(R, 1, "Grand Total", grandTotalStyle);
  setCell(R, 2, "", grandTotalStyle);
  setCell(R, 3, grandTotalSingle, { ...grandTotalStyle, t: "n" });
  setCell(R, 4, grandTotalOrder, { ...grandTotalStyle, t: "n" });
  R++;

  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: R - 1, c: 4 } });
  ws["!cols"] = [
    { wch: 5 },
    { wch: 20 },
    { wch: 55 },
    { wch: 14 },
    { wch: 14 },
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
      className="border rounded px-2 py-1 w-full text-sm text-center"
    />
  );
}

// ─── RawMaterialRow ───────────────────────────────────────────────────────────
function RawMaterialRow({
  material, bomQty, isEditing,
  onEdit, onSave, onCancel, onDelete, onChange, saving,
}) {
  const currentBomQty = parseFloat(bomQty) || 1;

  const handleSingleQtyChange = (field, value) => {
    const single = parseFloat(value) || 0;
    onChange("Single_Quantity", single);
    onChange("Round_Off", (single * currentBomQty).toFixed(2));
  };

  const handleOrderQtyChange = (value) => {
    const order = parseFloat(value) || 0;
    onChange("Round_Off", order);
    // onChange("Single_Quantity", (order / currentBomQty).toFixed(4));
  };

  return (
    <tr className="bg-white hover:bg-gray-50 transition-colors">
      <td className="border p-2 pl-12">
        <EditableField
          isEditing={isEditing}
          value={material.SKU}
          field="SKU"
          onChange={(f, v) => onChange(f, v)}
        />
      </td>
      <td className="border p-2 text-gray-600">
        <span>{material.Description || "—"}</span>
      </td>
      <td className="border p-2 text-gray-600">
        <span>{material.Po_Name || "—"}</span>
      </td>
      <td className="border p-2 text-center w-32">
        <EditableField
          isEditing={isEditing}
          value={material.Single_Quantity}
          field="Single_Quantity"
          type="number"
          onChange={handleSingleQtyChange}
        />
      </td>
      <td className="border p-2 text-center w-32">
        {isEditing ? (
          <input
            type="number"
            value={material.Round_Off ?? ""}
            onChange={(e) => handleOrderQtyChange(e.target.value)}
            className="border rounded px-2 py-1 w-full text-sm text-center font-medium text-gray-700 focus:outline-none focus:ring-1 focus:ring-blue-500"
          />
        ) : (
          <span className="font-medium text-gray-700">
            {parseFloat(material.Round_Off || 0).toFixed(2)}
          </span>
        )}
      </td>
      <td className="border p-2">
        <div className="flex gap-2 flex-wrap">
          {!isEditing && (
            <button onClick={onEdit} className="bg-blue-500 text-white px-2 py-1 rounded text-sm hover:bg-blue-600">
              Edit
            </button>
          )}
          {isEditing && (
            <>
              <button
                onClick={onSave}
                disabled={saving}
                className="bg-green-500 text-white px-2 py-1 rounded text-sm disabled:opacity-50 hover:bg-green-600"
              >
                {saving ? "Saving…" : "Save"}
              </button>
              <button
                onClick={onCancel}
                disabled={saving}
                className="bg-gray-400 text-white px-2 py-1 rounded text-sm hover:bg-gray-500"
              >
                Cancel
              </button>
            </>
          )}
          <button
            onClick={() => confirmDelete(onDelete)}
            className="bg-red-500 text-white px-2 py-1 rounded text-sm hover:bg-red-600"
          >
            Delete
          </button>
        </div>
      </td>
    </tr>
  );
}

// ─── BOMRow ───────────────────────────────────────────────────────────────────
function BOMRow({
  bom, isEditing, expanded,
  onToggle, onEdit, onSave, onCancel, onDelete, onAddRaw,
  onChange, saving,
}) {
  const ef = (field, type = "text") => (
    <EditableField isEditing={isEditing} value={bom[field]} field={field} type={type} onChange={onChange} />
  );

  return (
    <tr className="bg-gray-100 font-medium">
      <td className="border p-2">
        <div className="flex items-center gap-2">
          <button
            onClick={onToggle}
            className="w-6 h-6 bg-gray-300 hover:bg-gray-400 rounded flex items-center justify-center text-xs transition-colors"
          >
            {expanded ? "▼" : "▶"}
          </button>
          {ef("SKU")}
        </div>
        <div className="text-xs text-gray-500 mt-1 ml-8">{bom.rawMaterials?.length || 0} Raw Material(s)</div>
      </td>
      <td className="border p-2">{ef("Kit_Name")}</td>
      <td className="border p-2">{ef("Po_Name")}</td>
      <td className="border p-2 text-center text-gray-400">—</td>
      <td className="border p-2">
        {isEditing ? (
          <input
            type="number"
            value={bom.Qty ?? ""}
            min={1}
            onChange={(e) => onChange("Qty", e.target.value)}
            className="border rounded px-2 py-1 w-24 text-sm font-normal text-center"
            placeholder="Qty"
          />
        ) : (
          <span className="font-semibold text-gray-800 block text-center">{bom.Qty ?? "—"}</span>
        )}
      </td>
      <td className="border p-2">
        <div className="flex gap-2 flex-wrap">
          <button
            onClick={() => confirmRaw(onAddRaw)}
            className="bg-purple-500 text-white px-2 py-1 rounded text-sm hover:bg-purple-600"
          >
            + Add
          </button>
          {!isEditing && (
            <button onClick={onEdit} className="bg-blue-500 text-white px-2 py-1 rounded text-sm hover:bg-blue-600">
              Edit
            </button>
          )}
          {isEditing && (
            <>
              <button
                onClick={onSave}
                disabled={saving}
                className="bg-green-500 text-white px-2 py-1 rounded text-sm disabled:opacity-50 hover:bg-green-600"
              >
                {saving ? "Saving…" : "Save"}
              </button>
              <button
                onClick={onCancel}
                disabled={saving}
                className="bg-gray-400 text-white px-2 py-1 rounded text-sm hover:bg-gray-500"
              >
                Cancel
              </button>
            </>
          )}
          <button
            onClick={() => confirmDelete(onDelete)}
            className="bg-red-500 text-white px-2 py-1 rounded text-sm hover:bg-red-600"
          >
            Delete
          </button>
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

  // ── SKU Summary Calculation ─────────────────────────────────────────────────
  const skuSummary = useMemo(() => {
    const map = new Map();
    for (const bom of rows) {
      if (!bom.rawMaterials) continue;

      for (const mat of bom.rawMaterials) {
        const sku = (mat.SKU ?? "").trim();
        if (!sku) continue;
        const singleQty = parseFloat(mat.Single_Quantity) || 0;
        const orderQty = parseFloat(mat.Round_Off) || 0;

        if (map.has(sku)) {
          const entry = map.get(sku);
          entry.totalSingleQty += singleQty;
          entry.totalOrderQty += orderQty;
        } else {
          map.set(sku, {
            SKU: sku,
            Description: mat.Description || mat.Name || "",
            totalSingleQty: singleQty,
            totalOrderQty: orderQty,
          });
        }
      }
    }
    return Array.from(map.values()).sort((a, b) => a.SKU.localeCompare(b.SKU));
  }, [rows]);

  // ── Initial load ──────────────────────────────────────────────────────────
  useEffect(() => {

    const timer = setTimeout(() => {

      initialLoad();

    }, 2000);

    return () => clearTimeout(timer);

  }, []);

  async function initialLoad() {
    try {
      setLoading(true);
      const dealIdVal = bomID();
      if (!dealIdVal) return;

      const dealRes = await zoho().CRM.API.getRecord({ Entity: "Deals", RecordID: dealIdVal });
      const deal = dealRes?.data?.[0];
      if (deal) setDealData(deal);

      const bomList = (
        await fetchAllPages("BOM", `((Deals:equals:${dealIdVal}) and (BOM_Stage:equals:Initial BOM))`)
      ).filter((r) => r.BOM_Status !== "Cancelled");

      const finalBomList = (
        await fetchAllPages("BOM", `((Deals:equals:${dealIdVal}) and (BOM_Stage:equals:Final BOM))`)
      ).filter((r) => r.BOM_Status !== "Cancelled");

      setFinalBom(finalBomList);

      const rawPerBOM = await Promise.all(
        bomList.map((bom) => fetchAllPages("RawMaterials", `(BOM:equals:${bom.id})`))
      );

      setRows(
        bomList.map((bom, i) => ({
          ...bom,
          rawMaterials: (rawPerBOM[i] ?? []).map((r) => {
            const bQty = parseFloat(bom.Qty) || 1;
            const singleQty = r.Single_Quantity ?? r.Round_Off ?? 1;
            const orderQty = r.Round_Off ?? (singleQty * bQty);
            return {
              ...r,
              _bomId: bom.id,
              Description: r.Description ?? r.Name ?? "",
              Po_Name: r.Po_Name ?? "",
              Single_Quantity: singleQty,
              Round_Off: orderQty,
              SKU: r.SKU ?? "",
            };
          }),
        }))
      );
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
        if (rawIndex === null) {
          if (field === "Qty") {
            const nextQty = parseFloat(value) || 0;
            const updatedRaw = (bom.rawMaterials || []).map(mat => ({
              ...mat,
              Round_Off: nextQty > 0 ? ((parseFloat(mat.Single_Quantity) || 0) * nextQty).toFixed(2) : "0.00"
            }));
            return { ...bom, [field]: value, rawMaterials: updatedRaw };
          }
          return { ...bom, [field]: value };
        }
        return {
          ...bom,
          rawMaterials: bom.rawMaterials.map((mat, ri) =>
            ri !== rawIndex ? mat : { ...mat, [field]: value }
          ),
        };
      })
    );
  }

  // ── Save Logic Handler ────────────────────────────────────────────────────
  async function handleSave() {
    if (!editTarget) return;
    const { bomIndex, rawIndex } = editTarget;
    setSaving(true);
    try {
      if (rawIndex === null) {
        await saveBOM(bomIndex);
      } else {
        await saveRawMaterial(bomIndex, rawIndex);
      }
      backup.current = null;
      setEditTarget(null);
    } catch (err) {
      console.error(err);
      toast.error(err?.message || "Save failed");
    } finally {
      setSaving(false);
    }
  }

  // CRITICAL UPDATE: BOM Quantity update hote hi uske sabhi raw materials ko save karega
  async function saveBOM(bomIndex) {
    const row = rows[bomIndex];
    const payload = { SKU: row.SKU, Kit_Name: row.Kit_Name, Po_Name: row.Po_Name, Qty: parseInt(row.Qty) || 0 };
    const ownerId = users?.users?.[0]?.id || users?.[0]?.id;

    let targetId = row.id;

    if (row._isNew) {
      const res = await zoho().CRM.API.insertRecord({
        Entity: "BOM",
        APIData: {
          ...payload,
          Deals: { id: bomID() },
          Owner: ownerId ? { id: ownerId } : undefined,
          Account_Name: dealData?.Account_Name?.id ? { id: dealData.Account_Name.id } : undefined,
          BOM_Stage: "Initial BOM",
        },
        Trigger: ["workflow"],
      });
      targetId = res?.data?.[0]?.details?.id;
      if (!targetId) throw new Error("BOM create failed — no ID returned from Zoho");

      setRows((prev) =>
        prev.map((bom, bi) =>
          bi !== bomIndex ? bom : { ...bom, id: targetId, _isNew: false }
        )
      );
    } else {
      await zoho().CRM.API.updateRecord({
        Entity: "BOM",
        APIData: { id: row.id, ...payload },
        Trigger: ["workflow"],
      });
    }

    // Related Child Raw Materials Update Loop
    if (row.rawMaterials && row.rawMaterials.length > 0) {
      const promises = row.rawMaterials.map((mat) => {
        // Agar main BOM new tha toh dynamic ID pass karenge runtime par
        const rawPayload = {
          id: mat._isNew ? undefined : mat.id,
          SKU: mat.SKU,
          Description: mat.Description,
          Po_Name: mat.Po_Name,
          Single_Quantity: parseFloat(mat.Single_Quantity) || 0,
          Round_Off: parseFloat(mat.Round_Off) || 0, // Recalculated Order Qty
          BOM: { id: targetId }
        };

        if (mat._isNew) {
          return zoho().CRM.API.insertRecord({ Entity: "RawMaterials", APIData: rawPayload, Trigger: ["workflow"] });
        } else {
          return zoho().CRM.API.updateRecord({ Entity: "RawMaterials", APIData: rawPayload, Trigger: ["workflow"] });
        }
      });

      await Promise.all(promises);

      // State sync refresh agar kuch new records insert hue the nested items mein
      if (row._isNew) {
        initialLoad();
      }
    }

    toast.success("BOM and all related Raw Materials Saved successfully!");
  }

  async function saveRawMaterial(bomIndex, rawIndex) {
    const bomRow = rows[bomIndex];
    const row = bomRow.rawMaterials[rawIndex];
    const parentId = bomRow.id;

    const isDuplicate = bomRow.rawMaterials.some(
      (mat, ri) =>
        ri !== rawIndex &&
        mat.SKU?.trim().toLowerCase() === row.SKU?.trim().toLowerCase()
    );

    if (isDuplicate) {
      toast.error(
        `"${row.SKU}" is already added in this BOM.`
      );

      // Remove duplicate row
      setRows(prev =>
        prev.map((b, bi) =>
          bi !== bomIndex
            ? b
            : {
              ...b,
              rawMaterials: b.rawMaterials.filter(
                (_, ri) => ri !== rawIndex
              )
            }
        )
      );

      return;
    }

    const masterRes = await zoho().CRM.API.searchRecord({
      Entity: "Products",
      Type: "criteria",
      Query: `(Product_Code:equals:${encodeURIComponent(row.SKU)})`,
    });
    const product = masterRes?.data?.[0];
    if (!product) {
      toast.error("SKU not found in Product Master");
      return;
    }

    const description = product?.Description || product?.Product_Description || row.Description;

    const payload = {
      SKU: row.SKU,
      Description: description,
      Po_Name: row.Po_Name,
      Single_Quantity: parseFloat(row.Single_Quantity) || 0,
      Round_Off: parseFloat(row.Round_Off) || 0,
    };

    if (row._isNew) {
      const res = await zoho().CRM.API.insertRecord({
        Entity: "RawMaterials",
        APIData: { ...payload, BOM: { id: parentId }, Name: `${rawIndex + 1}`, S_No: `${rawIndex + 1}` },
        Trigger: ["workflow"],
      });
      const createdId = res?.data?.[0]?.details?.id;
      if (!createdId) throw new Error("Raw Material create failed");

      setRows((prev) =>
        prev.map((b, bi) =>
          bi !== bomIndex
            ? b
            : {
              ...b,
              rawMaterials: b.rawMaterials.map((mat, ri) =>
                ri !== rawIndex ? mat : { ...mat, id: createdId, Description: description, _isNew: false }
              ),
            }
        )
      );
      toast.success("Raw Material Created Successfully");
    } else {
      await zoho().CRM.API.updateRecord({
        Entity: "RawMaterials",
        APIData: { id: row.id, ...payload },
        Trigger: ["workflow"],
      });

      setRows((prev) =>
        prev.map((b, bi) =>
          bi !== bomIndex
            ? b
            : {
              ...b,
              rawMaterials: b.rawMaterials.map((mat, ri) =>
                ri !== rawIndex ? mat : { ...mat, Description: description }
              ),
            }
        )
      );
      toast.success("Raw Material Updated Successfully");
    }
  }

  // ── Add Triggers ──────────────────────────────────────────────────────────
  function addBOM() {
    const newBOM = {
      id: `_new_bom_${Date.now()}`,
      _isNew: true,
      SKU: "",
      Kit_Name: "",
      Po_Name: "",
      Qty: 1,
      rawMaterials: [],
    };
    backup.current = JSON.parse(JSON.stringify(rows));
    setRows((prev) => [...prev, newBOM]);
    setEditTarget({ bomIndex: rows.length, rawIndex: null });
  }

  function addRawMaterial(bomIndex) {
    const newMat = {
      id: `_new_raw_${Date.now()}`,
      _isNew: true,
      SKU: "",
      Description: "",
      Po_Name: "",
      Single_Quantity: 1,
      Round_Off: 1,
    };
    backup.current = JSON.parse(JSON.stringify(rows));
    setRows((prev) => prev.map((bom, bi) =>
      bi !== bomIndex ? bom : { ...bom, rawMaterials: [...(bom.rawMaterials || []), newMat] }
    ));

    const targetBOM = rows[bomIndex];
    setExpanded((e) => ({ ...e, [targetBOM.id]: true }));
    setEditTarget({ bomIndex, rawIndex: targetBOM.rawMaterials?.length || 0 });
  }

  // ── Delete Operations ─────────────────────────────────────────────────────
  async function removeBOM(bomIndex) {
    const bom = rows[bomIndex];
    try {
      if (!bom._isNew) await zoho().CRM.API.deleteRecord({ Entity: "BOM", RecordID: bom.id });
      setRows((prev) => prev.filter((_, bi) => bi !== bomIndex));
      toast.success("BOM deleted");
    } catch (err) {
      toast.error(err?.message || "Failed to delete BOM");
    }
  }

  async function removeRawMaterial(bomIndex, rawIndex) {
    const mat = rows[bomIndex].rawMaterials[rawIndex];
    try {
      if (!mat._isNew) await zoho().CRM.API.deleteRecord({ Entity: "RawMaterials", RecordID: mat.id });
      setRows((prev) =>
        prev.map((bom, bi) =>
          bi !== bomIndex
            ? bom
            : { ...bom, rawMaterials: bom.rawMaterials.filter((_, ri) => ri !== rawIndex) }
        )
      );
      toast.success("Raw Material deleted");
    } catch (err) {
      toast.error(err?.message || "Failed to delete Raw Material");
    }
  }

  if (loading) {
    return (
      <div className="h-full w-full flex justify-center items-center py-20">
        <div className="h-8 w-8 animate-spin rounded-full border-4 border-blue-600 border-t-transparent" />
      </div>
    );
  }

  //Reject Bom logic
  const rejectBom = async () => {
    // Update Blueprint Upload Bom 
    try {

      const BlueprintData = {
        blueprint: [
          {
            transition_id:
              "911594000013716281"
          }
        ]
      };

      const config = {
        Entity: "Deals",
        RecordID: bomID,
        BlueprintData
      };

      const response =
        await ZOHO.CRM.API
          .updateBluePrint(config);

      console.log(
        "Blueprint Updated",
        response
      );

      ZOHO.CRM.UI.Popup.closeReload()
        .then(function (data) {
          console.log(data)
        })
    } catch (err) {
      toast.error(err?.message || "Failed to Reject BOM");
    }
  }
  const updateReview = async () => {
    // Update Blueprint Upload Bom 
    try {
      if (finalBom.length > 0) {
        const BlueprintData = {
          blueprint: [
            {
              transition_id:
                "911594000013716280"
            }
          ]
        };

        const config = {
          Entity: "Deals",
          RecordID: bomID,
          BlueprintData
        };

        const response =
          await ZOHO.CRM.API
            .updateBluePrint(config);

        console.log(
          "Blueprint Updated",
          response
        );

        ZOHO.CRM.UI.Popup.closeReload()
          .then(function (data) {
            console.log(data)
          })
      } else {
        toast.error(
          "Please go to the Edit Table section and save the changes first."
        );
      }
    } catch (err) {
      toast.error(err?.message || "Failed to update blueprint");
    }
  }

  return (
    <div className="p-5">
      {/* Toolbar */}
      <div className="flex justify-between items-center">
        <div className="mb-4 flex items-center gap-3 flex-wrap">
          <button
            onClick={addBOM}
            className="bg-blue-600 text-white px-4 py-2 rounded hover:bg-blue-700 text-sm transition-colors"
          >
            + Add BOM
          </button>

          <button
            onClick={() => exportToExcel(rows, skuSummary, dealData)}
            disabled={rows.length === 0}
            className="flex items-center gap-2 bg-green-600 text-white px-4 py-2 rounded hover:bg-green-700 text-sm disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
              <polyline points="14 2 14 8 20 8" />
              <line x1="8" y1="13" x2="16" y2="13" />
              <line x1="8" y1="17" x2="16" y2="17" />
            </svg>
            Export Excel
          </button>
        </div>
        <div className="mb-4 flex items-center gap-3 flex-wrap">
          <button
            onClick={() => reviewConfirm(updateReview)}
            disabled={rows.length === 0}
            className="flex items-center gap-2 bg-green-600 text-white px-4 py-2 rounded hover:bg-green-700 text-sm disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            Reviewed
          </button>
          <button
            onClick={() => confirmRejectBom(rejectBom)}
            disabled={rows.length === 0}
            className="flex items-center gap-2 bg-red-600 text-white px-4 py-2 rounded hover:bg-red-700 text-sm disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            Reject BOM
          </button>
        </div>
      </div>

      {/* Main Grid View */}
      <table className="w-full border-collapse border text-sm text-left">
        <thead className="bg-gray-200 text-gray-700 font-semibold">
          <tr>
            <th className="border p-2 w-1/5">SKU</th>
            <th className="border p-2 w-1/4">Name of Component / Description</th>
            <th className="border p-2">Po Name</th>
            <th className="border p-2 text-center">Single Quantity</th>
            <th className="border p-2 text-center">Order Quantity</th>
            <th className="border p-2">Actions</th>
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
                {expanded[bom.id] && bom.rawMaterials?.map((mat, rawIndex) => {
                  const isRawEditing = editTarget?.bomIndex === bomIndex && editTarget?.rawIndex === rawIndex;
                  return (
                    <RawMaterialRow
                      key={mat.id}
                      material={mat}
                      bomQty={bom.Qty}
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
              <td colSpan={6} className="border p-8 text-center text-gray-400 font-medium">
                No BOMs found. Click "+ Add BOM" to create one.
              </td>
            </tr>
          )}
        </tbody>
      </table>

      {/* Merged SKU Totals Summary Component */}
      <SKUSummaryTable
        summary={skuSummary}
        dealData={dealData}
        dealId={bomID()}
        existingBOMId={finalBom && finalBom.length > 0 ? finalBom[0]?.id : null}
        users={users}
      />
    </div>
  );
}