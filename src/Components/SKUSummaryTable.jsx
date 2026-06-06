import { useState, useEffect, useCallback, useRef } from "react";

// ─────────────────────────────────────────────────────────────────────────────
// Zoho CRM — Inventory lookup
// ─────────────────────────────────────────────────────────────────────────────

async function fetchSKUFromInventory(sku) {
    if (typeof ZOHO === "undefined" || !ZOHO?.CRM) return null;
    try {
        const data = await ZOHO.CRM.API.searchRecord({
            Entity: "Products",
            Type: "criteria",
            Query: `(Product_Code:equals:${sku})`,
        });
        const record = data?.data?.[0];
        return record ? { lwpPrice: record?.Cost_Price ?? 0 } : null;
    } catch {
        return null;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Zoho CRM — BOM save / update
// ─────────────────────────────────────────────────────────────────────────────

function buildSubformRows(rows) {
    return rows.map((item) => ({
        Item_Code: item.SKU,
        Description: item.Description,
        Tagging_Qty: parseFloat(item.totalOrderQty || item.totalRoundOff || 0), // Standardized target order quantity
        Item_Rate: item.lwpPrice,
        Discount_Percent: item.discountPct,
        Discount_Price: item.dealerPrice,
        Anodizing_Factor: item.anodizingPct,
        Item_Total: item.fobPrice,
        Amount: item.amount,
    }));
}

async function saveToBOMModule({ rows, existingBOMId, dealData, users, dealId }) {
    if (typeof ZOHO === "undefined" || !ZOHO?.CRM) {
        console.warn("Zoho CRM SDK not available – dry run.");
        return { success: true, id: existingBOMId ?? "DRY_RUN_ID" };
    }

    const Raw_Material = buildSubformRows(rows);

    if (existingBOMId) {
        const result = await ZOHO.CRM.API.updateRecord({
            Entity: "BOM",
            APIData: { id: existingBOMId, BOM_Stage: "Final BOM", Raw_Material },
        });
        const ok = result?.data?.[0]?.status === "success";
        return { success: ok, id: existingBOMId };
    }

    const result = await ZOHO.CRM.API.insertRecord({
        Entity: "BOM",
        APIData: {
            BOM_Stage: "Final BOM",
            Deals: { id: dealId },
            Owner: { id: users?.users?.[0]?.id },
            Account_Name: { id: dealData?.Account_Name?.id },
            Raw_Material,
        },
    });
    const id = result?.data?.[0]?.details?.id;
    return { success: Boolean(id), id };
}

// ─────────────────────────────────────────────────────────────────────────────
// Merge saved BOM subform data back onto the summary rows
// ─────────────────────────────────────────────────────────────────────────────

function mergeBOMData(summary, bomLines = []) {
    const saved = Object.fromEntries(
        bomLines.map((line) => [line.Item_Code, line])
    );
    return summary.map((item) => {
        const s = saved[item.SKU];
        return initRow({
            ...item,
            lwpPrice: s ? (s.Item_Rate ?? 0) : (item.lwpPrice ?? 0),
            discountPct: s ? (s.Discount_Percent ?? 0) : (item.discountPct ?? 0),
            anodizingPct: s ? (s.Anodizing_Factor ?? 0) : (item.anodizingPct ?? 0),
        });
    });
}

// ─────────────────────────────────────────────────────────────────────────────
// Price calculation
// ─────────────────────────────────────────────────────────────────────────────

function calcRow(item) {
    const lwp = parseFloat(item.lwpPrice) || 0;
    const disc = parseFloat(item.discountPct) || 0;
    const anod = parseFloat(item.anodizingPct) || 0;
    // Fallback safely to totalOrderQty or totalRoundOff values coming from parent
    const qty = parseFloat(item.totalOrderQty || item.totalRoundOff || 0);

    const dealerPrice = lwp - (lwp * disc) / 100;
    const fobPrice = dealerPrice + (dealerPrice * anod) / 100;
    const amount = fobPrice * qty;

    return {
        ...item,
        dealerPrice: parseFloat(dealerPrice.toFixed(2)),
        fobPrice: parseFloat(fobPrice.toFixed(2)),
        amount: parseFloat(amount.toFixed(2)),
    };
}

function initRow(item) {
    return calcRow({
        ...item,
        lwpPrice: item.lwpPrice ?? 0,
        discountPct: item.discountPct ?? 0,
        dealerPrice: item.dealerPrice ?? 0,
        anodizingPct: item.anodizingPct ?? 0,
        fobPrice: item.fobPrice ?? 0,
        amount: item.amount ?? 0,
    });
}

// ─────────────────────────────────────────────────────────────────────────────
// UI atoms
// ─────────────────────────────────────────────────────────────────────────────

const STATUS_STYLE = {
    idle: "bg-amber-100  text-amber-700  border-amber-300",
    loading: "bg-gray-100   text-gray-500   border-gray-300  animate-pulse",
    saving: "bg-blue-100   text-blue-700   border-blue-300  animate-pulse",
    saved: "bg-emerald-100 text-emerald-700 border-emerald-300",
    error: "bg-red-100    text-red-700    border-red-300",
};
const STATUS_LABEL = {
    idle: "Unsaved",
    loading: "Loading…",
    saving: "Saving…",
    saved: "Saved ✓",
    error: "Error ✕",
};

function StatusPill({ status }) {
    return (
        <span className={`text-xs font-medium px-2 py-0.5 rounded-full border ${STATUS_STYLE[status] ?? STATUS_STYLE.idle}`}>
            {STATUS_LABEL[status] ?? STATUS_LABEL.idle}
        </span>
    );
}

function NumberInput({ value, onChange, error }) {
    return (
        <div className="flex flex-col items-end gap-0.5">
            <input
                type="number"
                value={value}
                onChange={(e) => onChange(e.target.value)}
                className={`w-full text-right rounded px-1 py-0.5 text-sm font-mono focus:outline-none focus:ring-1 ${error
                    ? "bg-red-50 border border-red-400 focus:ring-red-400 text-red-700"
                    : "bg-amber-50 border border-amber-300 focus:ring-amber-400"
                    }`}
                style={{ minWidth: 72 }}
            />
            {error && (
                <span className="text-red-500 text-[10px] leading-tight whitespace-nowrap">{error}</span>
            )}
        </div>
    );
}

const IconRefresh = () => (
    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
            d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
    </svg>
);
const IconCheck = () => (
    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
    </svg>
);
const IconEdit = () => (
    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
            d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
    </svg>
);
const IconClose = () => (
    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
    </svg>
);
const IconInfo = () => (
    <svg className="w-4 h-4 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
            d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
);

// ─────────────────────────────────────────────────────────────────────────────
// Main component
// ─────────────────────────────────────────────────────────────────────────────

export default function SKUSummaryTable({ summary, dealData, dealId, users, existingBOMId = null }) {
    const [status, setStatus] = useState("loading");
    const [editing, setEditing] = useState(false);
    const [loadingLWP, setLoadingLWP] = useState(false);
    const [bomId, setBomId] = useState(existingBOMId);
    const [rowErrors, setRowErrors] = useState({});
    const [rows, setRows] = useState(() => summary.map(initRow));
    const [savedRows, setSavedRows] = useState(() => summary.map(initRow));

    const mountDone = useRef(false);
    const savedPricingRef = useRef({});

    if (!summary?.length) return null;

    // ── On mount: load saved BOM data from CRM ────────────────────────────────
    useEffect(() => {
        let cancelled = false;

        async function load() {
            setStatus("loading");
            let bomRecord = null;
            let resolvedBomId = existingBOMId;

            if (resolvedBomId) {
                try {
                    const data = await ZOHO.CRM.API.getRecord({
                        Entity: "BOM",
                        RecordID: resolvedBomId,
                    });
                    bomRecord = data?.data?.[0] ?? null;
                } catch { /* not found — treat as new */ }
            }

            if (cancelled) return;

            if (bomRecord) {
                const bomLines = bomRecord.Raw_Material ?? [];
                const hydrated = mergeBOMData(summary, bomLines);

                savedPricingRef.current = Object.fromEntries(
                    hydrated.map((r) => [r.SKU, {
                        lwpPrice: r.lwpPrice,
                        discountPct: r.discountPct,
                        anodizingPct: r.anodizingPct,
                    }])
                );

                setRows(hydrated);
                setSavedRows(hydrated);
                setBomId(resolvedBomId);
                setStatus("saved");
            } else {
                const fresh = summary.map(initRow);
                setRows(fresh);
                setSavedRows(fresh);
                setStatus("idle");
            }

            mountDone.current = true;
        }

        if (typeof ZOHO !== "undefined" && ZOHO?.CRM) {
            load();
        } else {
            mountDone.current = true;
            setStatus("idle");
        }

        return () => { cancelled = true; };
    }, []);

    // ── When summary changes after mount → auto-save to BOM if bomId exists ──────
    useEffect(() => {
        if (!mountDone.current) return;
        if (!summary?.length) return;
        if (status === "loading") return;

        const fresh = summary.map((item) => {
            const saved = savedPricingRef.current[item.SKU];
            return initRow({
                ...item,
                lwpPrice: saved?.lwpPrice ?? 0,
                discountPct: saved?.discountPct ?? 0,
                anodizingPct: saved?.anodizingPct ?? 0,
            });
        });

        setRows(fresh);
        setSavedRows(fresh);
        setEditing(false);

        if (bomId) {
            setStatus("saving");
            const Raw_Material = buildSubformRows(fresh);

            ZOHO.CRM.API.updateRecord({
                Entity: "BOM",
                APIData: { id: bomId, BOM_Stage: "Final BOM", Raw_Material },
            })
                .then((result) => {
                    const ok = result?.data?.[0]?.status === "success";
                    if (ok) {
                        savedPricingRef.current = Object.fromEntries(
                            fresh.map((r) => [r.SKU, {
                                lwpPrice: r.lwpPrice,
                                discountPct: r.discountPct,
                                anodizingPct: r.anodizingPct,
                            }])
                        );
                        setStatus("saved");
                    } else {
                        setStatus("error");
                    }
                })
                .catch((err) => {
                    console.error("Auto-save BOM error:", err);
                    setStatus("error");
                });
        } else {
            setStatus("idle");
        }

    }, [summary]);

    // ── Edit / Cancel ─────────────────────────────────────────────────────────
    function handleEditToggle() {
        if (editing) setRows(savedRows);
        setEditing((v) => !v);
    }

    // ── Fetch LWP from Inventory ──────────────────────────────────────────────
    const handleFetchLWP = useCallback(async () => {
        setLoadingLWP(true);
        const updated = await Promise.all(
            rows.map(async (row) => {
                if (row.lwpPrice && Number(row.lwpPrice) > 0) return row;
                const inv = await fetchSKUFromInventory(row.SKU);
                return calcRow({ ...row, lwpPrice: inv?.lwpPrice ?? row.lwpPrice });
            })
        );
        setRows(updated);
        setLoadingLWP(false);
    }, [rows]);

    // ── Cell change ───────────────────────────────────────────────────────────
    function handleChange(idx, field, value) {
        const isPctField = field === "discountPct" || field === "anodizingPct";
        const pct = parseFloat(value);
        const errorKey = `${idx}_${field}`;
        const label = field === "discountPct" ? "Discount" : "Anodizing factor";

        if (isPctField && !isNaN(pct) && pct > 100) {
            setRowErrors((prev) => ({ ...prev, [errorKey]: `${label} cannot exceed 100%` }));
            return;
        }
        setRowErrors((prev) => { const n = { ...prev }; delete n[errorKey]; return n; });
        setRows((prev) => {
            const next = [...prev];
            next[idx] = calcRow({ ...next[idx], [field]: value });
            return next;
        });
    }

    // ── Save / Update BOM ─────────────────────────────────────────────────────
    async function handleSave() {
        setStatus("saving");
        try {
            const { success, id } = await saveToBOMModule({
                rows,
                existingBOMId: bomId,
                dealData,
                dealId,
                users,
            });
            if (success) {
                savedPricingRef.current = Object.fromEntries(
                    rows.map((r) => [r.SKU, {
                        lwpPrice: r.lwpPrice,
                        discountPct: r.discountPct,
                        anodizingPct: r.anodizingPct,
                    }])
                );
                setBomId(id);
                setSavedRows(rows);
                setStatus("saved");
                setEditing(false);
            } else {
                setStatus("error");
            }
        } catch (err) {
            console.error("BOM save error:", err);
            setStatus("error");
        }
    }

    // ── Totals Calculation ────────────────────────────────────────────────────
    const totalQty = rows.reduce((s, r) => s + (parseFloat(r.totalOrderQty || r.totalRoundOff || 0)), 0);
    const totalAmt = rows.reduce((s, r) => s + (parseFloat(r.amount) || 0), 0);

    const fmt = (n) =>
        typeof n === "number"
            ? n.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
            : "—";

    const isLoading = status === "loading";

    return (
        <div className="mt-8 font-sans">
            {/* Header */}
            <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
                <div className="flex items-center gap-2">
                    <h2 className="text-base font-semibold text-gray-700">Consolidated Data</h2>
                    <span className="text-xs text-gray-400 bg-gray-100 px-2 py-0.5 rounded-full border border-gray-200">
                        Final BOM
                    </span>
                    <StatusPill status={status} />
                    {bomId && (
                        <span className="text-xs text-gray-400 font-mono">
                            BOM: <span className="text-indigo-500">{bomId}</span>
                        </span>
                    )}
                </div>

                <div className="flex items-center gap-2">
                    {status === "idle" && !isLoading && (
                        <button
                            onClick={handleSave}
                            disabled={status === "saving"}
                            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-lg bg-emerald-600 text-white hover:bg-emerald-700 transition disabled:opacity-60 shadow-sm"
                        >
                            <IconCheck />
                            {bomId ? "Update BOM" : "Save as Final BOM"}
                        </button>
                    )}

                    {editing && !isLoading && (
                        <>
                            <button
                                onClick={handleFetchLWP}
                                disabled={loadingLWP}
                                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-lg border border-indigo-300 bg-indigo-50 text-indigo-700 hover:bg-indigo-100 transition disabled:opacity-50"
                            >
                                {loadingLWP
                                    ? <span className="w-3 h-3 border-2 border-indigo-400 border-t-transparent rounded-full animate-spin" />
                                    : <IconRefresh />
                                }
                                Fetch LWP from Inventory
                            </button>

                            <button
                                onClick={handleSave}
                                disabled={status === "saving"}
                                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-lg bg-emerald-600 text-white hover:bg-emerald-700 transition disabled:opacity-60 shadow-sm"
                            >
                                <IconCheck />
                                {bomId ? "Update BOM" : "Save as Final BOM"}
                            </button>
                        </>
                    )}

                    <button
                        onClick={handleEditToggle}
                        disabled={isLoading}
                        className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-lg border transition shadow-sm disabled:opacity-40 ${editing
                            ? "border-red-300 bg-red-50 text-red-700 hover:bg-red-100"
                            : "border-blue-300 bg-blue-50 text-blue-700 hover:bg-blue-100"
                            }`}
                    >
                        {editing ? <><IconClose /> Cancel</> : <><IconEdit /> Edit Table</>}
                    </button>
                </div>
            </div>

            {/* Banners */}
            {status === "idle" && !editing && (
                <div className="mb-2 px-3 py-2 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-700 flex items-center gap-2">
                    <IconInfo />
                    Summary updated — click <strong className="mx-0.5">{bomId ? "Update BOM" : "Save as Final BOM"}</strong> to save the latest changes.
                </div>
            )}

            {editing && (
                <div className="mb-2 px-3 py-2 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-700 flex items-center gap-2">
                    <IconInfo />
                    Edit <strong className="mx-0.5">Discount %</strong> and{" "}
                    <strong className="mx-0.5">Anodizing Factor %</strong> — Dealer Price, FOB Price
                    and Amount update automatically.
                </div>
            )}

            {isLoading && (
                <div className="mb-2 px-3 py-2 bg-gray-50 border border-gray-200 rounded-lg text-xs text-gray-500 flex items-center gap-2 animate-pulse">
                    <span className="w-3 h-3 border-2 border-gray-400 border-t-transparent rounded-full animate-spin" />
                    Loading saved BOM data from CRM…
                </div>
            )}

            {/* Table */}
            <div className={`overflow-x-auto rounded-xl border border-gray-200 shadow-sm transition-opacity ${isLoading ? "opacity-50 pointer-events-none" : "opacity-100"}`}>
                <table className="w-full border-collapse text-sm min-w-[900px]">
                    <thead>
                        <tr className="bg-gradient-to-r from-blue-600 to-blue-700 text-white">
                            {["#", "SKU", "Description", "Order Qty", "LWP Price", "Discount %", "Dealer Price", "Anodizing %", "FOB Price", "Amount"].map((h, i) => (
                                <th
                                    key={h}
                                    className={`px-3 py-2.5 font-semibold text-xs tracking-wide ${i <= 2 ? "text-left" : "text-right"
                                        } ${editing && (h === "Discount %" || h === "Anodizing %") ? "bg-amber-500" : ""}`}
                                >
                                    {h}
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {rows.map((item, i) => (
                            <tr
                                key={item.SKU ?? i}
                                className={`transition-colors hover:bg-blue-50 ${i % 2 === 0 ? "bg-white" : "bg-slate-50"}`}
                            >
                                <td className="px-3 py-2 text-gray-400 text-xs">{i + 1}</td>
                                <td className="px-3 py-2 font-mono text-blue-700 font-medium">{item.SKU || "—"}</td>
                                <td className="px-3 py-2 text-gray-700 max-w-xs truncate" title={item.Description}>
                                    {item.Description || "—"}
                                </td>
                                {/* Modified Cell: Render specific order quantity / round off value safely */}
                                <td className="px-3 py-2 text-right font-semibold text-gray-800 font-mono">
                                    {parseFloat(item.totalOrderQty || item.totalRoundOff || 0)}
                                </td>
                                <td className="px-3 py-2 text-right font-mono text-gray-700">
                                    {editing
                                        ? <NumberInput value={item?.lwpPrice} onChange={(v) => handleChange(i, "lwpPrice", v)} />
                                        : item?.lwpPrice
                                    }
                                </td>
                                <td className={`px-3 py-2 text-right font-mono ${editing ? "bg-amber-50" : ""}`}>
                                    {editing
                                        ? <NumberInput value={item.discountPct} onChange={(v) => handleChange(i, "discountPct", v)} error={rowErrors[`${i}_discountPct`]} />
                                        : <span className="text-orange-600">{item.discountPct}%</span>
                                    }
                                </td>
                                <td className="px-3 py-2 text-right font-mono">
                                    <span className="text-emerald-700 font-medium">{fmt(item.dealerPrice)}</span>
                                </td>
                                <td className={`px-3 py-2 text-right font-mono ${editing ? "bg-amber-50" : ""}`}>
                                    {editing
                                        ? <NumberInput value={item.anodizingPct} onChange={(v) => handleChange(i, "anodizingPct", v)} error={rowErrors[`${i}_anodizingPct`]} />
                                        : <span className="text-purple-600">{item.anodizingPct}%</span>
                                    }
                                </td>
                                <td className="px-3 py-2 text-right font-mono font-semibold text-blue-700">
                                    {fmt(item.fobPrice)}
                                </td>
                                <td className="px-3 py-2 text-right font-mono font-bold text-gray-900">
                                    {fmt(item.amount)}
                                </td>
                            </tr>
                        ))}
                    </tbody>
                    <tfoot>
                        <tr className="bg-gradient-to-r from-gray-100 to-blue-50 border-t-2 border-blue-200">
                            <td className="px-3 py-2.5 font-bold text-gray-700 text-sm" colSpan={3}>Grand Total</td>
                            <td className="px-3 py-2.5 text-right font-bold text-gray-800 font-mono">{totalQty}</td>
                            <td colSpan={5} />
                            <td className="px-3 py-2.5 text-right font-bold text-blue-700 font-mono text-base">
                                ₹{fmt(totalAmt)}
                            </td>
                        </tr>
                    </tfoot>
                </table>
            </div>

            {/* Formula legend */}
            <div className="mt-3 flex flex-wrap gap-x-4 gap-y-1 text-xs text-gray-400">
                <span><span className="text-orange-500 font-semibold">Dealer Price</span> = LWP − (LWP × Discount%)</span>
                <span className="text-gray-300">|</span>
                <span><span className="text-blue-500 font-semibold">FOB Price</span> = Dealer Price + (Dealer Price × Anodizing%)</span>
                <span className="text-gray-300">|</span>
                <span><span className="text-gray-600 font-semibold">Amount</span> = FOB Price × Qty</span>
            </div>
        </div>
    );
}