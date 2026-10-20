"use client";

import { useEffect, useMemo, useRef, useState, type ChangeEvent, type FormEvent } from "react";
import { ArrowDownToLine, Check, ChevronDown, ChevronLeft, ChevronRight, ClipboardList, Fan, MapPin, Pencil, Plus, Search, Trash2, Upload, X } from "lucide-react";
import { assetFields, demoAssets, type Asset, type AssetField } from "@/lib/assets";
import type { ImportedAssetRow } from "@/lib/asset-excel";
import { deleteAssetFromFirestore, deleteWorkbookPageFromFirestore, firebaseConfigured, initializeAssetsInFirestore, saveAssetToFirestore, saveAssetsToFirestore, saveIndexEntriesToFirestore, saveTermDefinitionsToFirestore, saveWorkbookPageToFirestore, watchAssetsInFirestore, watchIndexEntriesInFirestore, watchTermDefinitionsInFirestore, watchWorkbookPagesInFirestore } from "@/lib/firebase";
import { defaultSheets, indexEntries, termDefinitions, type IndexEntry, type TermDefinition, type WorkbookSheet } from "@/lib/sheets";

const storageKey = "mepf-assets-v1";
const sheetsStorageKey = "mepf-workbook-sheets-v1";
const indexStorageKey = "mepf-index-entries-v1";
const termsStorageKey = "mepf-term-definitions-v1";
const pageSize = 12;

const fieldLabels: Record<AssetField, string> = {
  equipmentGroup: "Equipment group",
  equipmentType: "Equipment type",
  equipmentTag: "Equipment tag",
  equipmentLocation: "Equipment location",
  servicedArea: "Serviced area",
  sequenceNumber: "Sequence number",
  equipmentCode: "Equipment code",
  brand: "Brand",
  capacity: "Capacity",
  hotelAssetCode: "Hotel asset code",
  installationYear: "Installation year",
  normalLife: "Life expectancy · normal",
  presentCondition: "Life expectancy · present condition",
  adjustedLife: "Life expectancy · adjusted",
  replacementYear: "Replacement year",
  zoning: "Zoning",
  phasing: "Phasing",
  remarks: "Remarks",
};

function blankAsset(): Asset {
  return {
    id: "",
    equipmentGroup: "",
    equipmentType: "",
    equipmentTag: "",
    equipmentLocation: "",
    servicedArea: "",
    sequenceNumber: "",
    equipmentCode: "",
    brand: "",
    capacity: "",
    hotelAssetCode: "",
    installationYear: "",
    normalLife: "",
    presentCondition: "",
    adjustedLife: "",
    replacementYear: "",
    zoning: "",
    phasing: "",
    remarks: "",
    remarksColor: "",
  };
}

function csvValue(value: string) {
  return `"${value.replaceAll('"', '""')}"`;
}

function errorDetails(error: unknown) {
  if (error && typeof error === "object") {
    const detail = error as { code?: unknown; message?: unknown };
    const code = typeof detail.code === "string" ? detail.code : "Error";
    const message = typeof detail.message === "string" ? detail.message : "Unknown error";
    return `${code}: ${message}`;
  }
  return error instanceof Error ? error.message : "Unknown error";
}

export default function Home() {
  const [assets, setAssets] = useState<Asset[]>(demoAssets);
  const [sheets, setSheets] = useState<WorkbookSheet[]>(defaultSheets);
  const [indexRows, setIndexRows] = useState<IndexEntry[]>(indexEntries);
  const [definitionRows, setDefinitionRows] = useState<TermDefinition[]>(termDefinitions);
  const [activeSheetId, setActiveSheetId] = useState("index");
  const [addingPage, setAddingPage] = useState(false);
  const [newPageName, setNewPageName] = useState("");
  const [search, setSearch] = useState("");
  const [locationFilter, setLocationFilter] = useState("All locations");
  const [phaseFilter, setPhaseFilter] = useState("All phases");
  const [page, setPage] = useState(1);
  const [editingAsset, setEditingAsset] = useState<Asset | null>(null);
  const [editingIndexRow, setEditingIndexRow] = useState<IndexEntry | null>(null);
  const [editingDefinition, setEditingDefinition] = useState<TermDefinition | null>(null);
  const [editingDefinitionOriginalTerm, setEditingDefinitionOriginalTerm] = useState<string | null>(null);
  const [pendingImport, setPendingImport] = useState<{ rows: ImportedAssetRow[]; group: string; groupLabel: string } | null>(null);
  const [referenceFormError, setReferenceFormError] = useState("");
  const [assetsInitialized, setAssetsInitialized] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [notice, setNotice] = useState("");
  const [cloudStatus, setCloudStatus] = useState(firebaseConfigured ? "connecting" : "local");
  const importFileInput = useRef<HTMLInputElement>(null);

  useEffect(() => {
    let active = true;
    let unsubscribePages: (() => void) | undefined;
    let unsubscribeIndex: (() => void) | undefined;
    let unsubscribeDefinitions: (() => void) | undefined;

    if (firebaseConfigured) {
      unsubscribePages = watchWorkbookPagesInFirestore((storedPages) => {
        if (!active) return;
        setSheets([...defaultSheets, ...storedPages]);
        window.localStorage.setItem(sheetsStorageKey, JSON.stringify(storedPages));
      }, (error) => {
        if (!active) return;
        setCloudStatus("error");
        setNotice(`Workbook page sync failed: ${error.message}`);
      });
      unsubscribeIndex = watchIndexEntriesInFirestore((storedEntries) => {
        if (!active) return;
        setIndexRows(storedEntries);
        window.localStorage.setItem(indexStorageKey, JSON.stringify(storedEntries));
      }, (error) => {
        if (!active) return;
        setCloudStatus("error");
        setNotice(`Index sync failed: ${error.message}`);
      });
      unsubscribeDefinitions = watchTermDefinitionsInFirestore((storedDefinitions) => {
        if (!active) return;
        setDefinitionRows(storedDefinitions);
        window.localStorage.setItem(termsStorageKey, JSON.stringify(storedDefinitions));
      }, (error) => {
        if (!active) return;
        setCloudStatus("error");
        setNotice(`Definition sync failed: ${error.message}`);
      });

      void initializeAssetsInFirestore().then(() => {
        if (!active) return;
        setAssetsInitialized(true);
      }).catch((error: unknown) => {
        if (!active) return;
        setCloudStatus("error");
        const message = error instanceof Error ? error.message : "Firestore could not be reached.";
        setNotice(`Firestore sync failed: ${message}`);
      });
    } else {
      const storedAssets = window.localStorage.getItem(storageKey);
      if (storedAssets) {
        const parsedAssets = JSON.parse(storedAssets) as Asset[];
        queueMicrotask(() => {
          if (active) setAssets(parsedAssets);
        });
      }
    }

    return () => {
      active = false;
      unsubscribePages?.();
      unsubscribeIndex?.();
      unsubscribeDefinitions?.();
    };
  }, []);

  useEffect(() => {
    if (!firebaseConfigured || !assetsInitialized) return;
    const groups = sheets.filter((sheet) => sheet.kind === "group").map((sheet) => sheet.group ?? sheet.label);
    return watchAssetsInFirestore((storedAssets) => {
      setAssets(storedAssets);
      window.localStorage.setItem(storageKey, JSON.stringify(storedAssets));
      setCloudStatus("connected");
    }, (error) => {
      setCloudStatus("error");
      setNotice(`Firestore sync failed: ${error.message}`);
    }, groups);
  }, [assetsInitialized, sheets]);

  useEffect(() => {
    if (firebaseConfigured) return;
    const savedSheets = window.localStorage.getItem(sheetsStorageKey);
    if (savedSheets) {
      const customSheets = JSON.parse(savedSheets) as WorkbookSheet[];
      queueMicrotask(() => setSheets([...defaultSheets, ...customSheets]));
    }
  }, []);

  useEffect(() => {
    if (firebaseConfigured) return;
    const savedIndex = window.localStorage.getItem(indexStorageKey);
    const savedDefinitions = window.localStorage.getItem(termsStorageKey);
    if (savedIndex) {
      const parsedIndex = JSON.parse(savedIndex) as IndexEntry[];
      queueMicrotask(() => setIndexRows(parsedIndex));
    }
    if (savedDefinitions) {
      const parsedDefinitions = JSON.parse(savedDefinitions) as TermDefinition[];
      queueMicrotask(() => setDefinitionRows(parsedDefinitions));
    }
  }, []);

  const activeSheet = sheets.find((sheet) => sheet.id === activeSheetId) ?? defaultSheets[0];

  const locations = useMemo(
    () => [...new Set(assets.map((asset) => asset.equipmentLocation).filter(Boolean))].sort(),
    [assets],
  );
  const phases = useMemo(
    () => [...new Set(assets.map((asset) => asset.phasing).filter(Boolean))].sort(),
    [assets],
  );
  const filteredAssets = useMemo(() => {
    const query = search.trim().toLowerCase();
    const sheetAssets = activeSheet.kind === "group"
      ? assets.filter((asset) => asset.equipmentGroup.toLowerCase() === activeSheet.group?.toLowerCase())
      : assets;
    return sheetAssets.filter((asset) => {
      const matchesSearch = !query || assetFields.some((field) => asset[field].toLowerCase().includes(query));
      const matchesLocation = locationFilter === "All locations" || asset.equipmentLocation === locationFilter;
      const matchesPhase = phaseFilter === "All phases" || asset.phasing === phaseFilter;
      return matchesSearch && matchesLocation && matchesPhase;
    });
  }, [activeSheet, assets, locationFilter, phaseFilter, search]);
  const pageCount = Math.max(1, Math.ceil(filteredAssets.length / pageSize));
  const currentPage = Math.min(page, pageCount);
  const visibleAssets = filteredAssets.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  const firstVisibleRow = filteredAssets.length === 0 ? 0 : (currentPage - 1) * pageSize + 1;
  const lastVisibleRow = Math.min(currentPage * pageSize, filteredAssets.length);

  function persistLocally(nextAssets: Asset[]) {
    window.localStorage.setItem(storageKey, JSON.stringify(nextAssets));
  }

  async function saveIndexRows(nextRows: IndexEntry[]) {
    if (firebaseConfigured) await saveIndexEntriesToFirestore(nextRows);
    window.localStorage.setItem(indexStorageKey, JSON.stringify(nextRows));
    setIndexRows(nextRows);
  }

  async function saveDefinitions(nextRows: TermDefinition[]) {
    if (firebaseConfigured) await saveTermDefinitionsToFirestore(nextRows);
    window.localStorage.setItem(termsStorageKey, JSON.stringify(nextRows));
    setDefinitionRows(nextRows);
  }

  async function handleIndexRowSave(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!editingIndexRow) return;
    const number = editingIndexRow.number.trim();
    const equipmentGroup = editingIndexRow.equipmentGroup.trim();
    const description = editingIndexRow.description.trim();
    if (!description) {
      setReferenceFormError("Description is required.");
      return;
    }
    if (number && !/^\d+$/.test(number)) {
      setReferenceFormError("The index number must contain digits only.");
      return;
    }
    const savedRow: IndexEntry = {
      ...editingIndexRow,
      id: editingIndexRow.id || `index-${crypto.randomUUID()}`,
      number,
      equipmentGroup,
      description,
    };
    const nextRows = indexRows.some((row) => row.id === savedRow.id)
      ? indexRows.map((row) => row.id === savedRow.id ? savedRow : row)
      : [...indexRows, savedRow];
    setIsSaving(true);
    try {
      await saveIndexRows(nextRows);
      setEditingIndexRow(null);
      setNotice(firebaseConfigured ? "Index entry saved to Firestore." : "Index entry saved on this device.");
    } catch (error) {
      setReferenceFormError(`Index save failed: ${errorDetails(error)}`);
    } finally {
      setIsSaving(false);
    }
  }

  async function handleDefinitionSave(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!editingDefinition) return;
    const term = editingDefinition.term.trim();
    const definition = editingDefinition.definition.trim();
    if (!term || !definition) {
      setReferenceFormError("Both term and definition are required.");
      return;
    }
    if (definitionRows.some((row) =>
      row.term.toLowerCase() === term.toLowerCase()
      && row.term !== editingDefinitionOriginalTerm,
    )) {
      setReferenceFormError("A definition for that term already exists.");
      return;
    }
    const savedDefinition = { term, definition };
    const nextRows = editingDefinitionOriginalTerm !== null
      ? definitionRows.map((row) => row.term === editingDefinitionOriginalTerm ? savedDefinition : row)
      : [...definitionRows, savedDefinition];
    setIsSaving(true);
    try {
      await saveDefinitions(nextRows);
      setEditingDefinition(null);
      setEditingDefinitionOriginalTerm(null);
      setNotice(firebaseConfigured ? "Definition saved to Firestore." : "Definition saved on this device.");
    } catch (error) {
      setReferenceFormError(`Definition save failed: ${errorDetails(error)}`);
    } finally {
      setIsSaving(false);
    }
  }

  async function deleteIndexRow(row: IndexEntry) {
    if (!window.confirm(`Delete the “${row.description}” index entry?`)) return;
    setIsSaving(true);
    try {
      await saveIndexRows(indexRows.filter((item) => item.id !== row.id));
      setNotice(firebaseConfigured ? "Index entry deleted from Firestore." : "Index entry deleted.");
    } catch (error) {
      setNotice(`Index delete failed: ${errorDetails(error)}`);
    } finally {
      setIsSaving(false);
    }
  }

  async function deleteDefinition(definition: TermDefinition) {
    if (!window.confirm(`Delete the “${definition.term}” definition?`)) return;
    setIsSaving(true);
    try {
      await saveDefinitions(definitionRows.filter((item) => item.term !== definition.term));
      setNotice(firebaseConfigured ? "Definition deleted from Firestore." : "Definition deleted.");
    } catch (error) {
      setNotice(`Definition delete failed: ${errorDetails(error)}`);
    } finally {
      setIsSaving(false);
    }
  }

  async function addWorkbookPage(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const label = newPageName.trim();
    if (!label) return;
    if (sheets.some((sheet) => sheet.label.toLowerCase() === label.toLowerCase())) {
      setNotice("A page with that name already exists.");
      return;
    }
    const sheet: WorkbookSheet = { id: `custom-${crypto.randomUUID()}`, label, kind: "group", group: label };
    const nextSheets = [...sheets, sheet];
    setIsSaving(true);
    try {
      if (firebaseConfigured) await saveWorkbookPageToFirestore(sheet);
      setSheets(nextSheets);
      window.localStorage.setItem(sheetsStorageKey, JSON.stringify(nextSheets.filter((item) => !defaultSheets.some((initial) => initial.id === item.id))));
      setActiveSheetId(sheet.id);
      setPage(1);
      setNewPageName("");
      setAddingPage(false);
      setNotice(firebaseConfigured ? "Page saved to Firestore." : "Page saved on this device.");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Firestore could not save the page.";
      setNotice(`Page save failed: ${message}`);
    } finally {
      setIsSaving(false);
    }
  }

  async function deleteWorkbookPage(sheet: WorkbookSheet) {
    if (!window.confirm(`Delete the “${sheet.label}” page? Its equipment records will remain in the workbook.`)) return;
    const customSheets = sheets.filter((item) => item.id !== sheet.id && !defaultSheets.some((initial) => initial.id === item.id));
    setIsSaving(true);
    try {
      if (firebaseConfigured) await deleteWorkbookPageFromFirestore(sheet.id);
      setSheets([...defaultSheets, ...customSheets]);
      window.localStorage.setItem(sheetsStorageKey, JSON.stringify(customSheets));
      if (activeSheetId === sheet.id) {
        setActiveSheetId("index");
        setPage(1);
        setSearch("");
        setLocationFilter("All locations");
        setPhaseFilter("All phases");
      }
      setNotice(firebaseConfigured ? "Page deleted from Firestore." : "Page deleted from this device.");
    } catch (error) {
      setNotice(`Page delete failed: ${errorDetails(error)}`);
    } finally {
      setIsSaving(false);
    }
  }

  function startAddingAsset() {
    const asset = blankAsset();
    if (activeSheet.kind === "group") asset.equipmentGroup = activeSheet.group ?? activeSheet.label;
    if (activeSheet.kind === "life") asset.equipmentGroup = "Unassigned";
    setEditingAsset(asset);
  }

  function exportCsv() {
    const header = assetFields.map((field) => csvValue(fieldLabels[field])).join(",");
    const rows = filteredAssets.map((asset) => assetFields.map((field) => csvValue(asset[field])).join(","));
    const blob = new Blob([[header, ...rows].join("\r\n")], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "mepf-asset-register.csv";
    link.click();
    URL.revokeObjectURL(url);
  }

  async function downloadExcelTemplate() {
    const group = activeSheet.group ?? activeSheet.label;
    const groupAssets = assets.filter((asset) => asset.equipmentGroup.toLowerCase() === group.toLowerCase());
    try {
      const { downloadAssetTemplate } = await import("@/lib/asset-excel");
      await downloadAssetTemplate(group, groupAssets);
      setNotice(`Excel template downloaded for ${activeSheet.label}.`);
    } catch (error) {
      setNotice(`Could not create the Excel template: ${errorDetails(error)}`);
    }
  }

  async function handleExcelUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    const group = activeSheet.group ?? activeSheet.label;
    const groupLabel = activeSheet.label;
    const currentGroupAssets = assets.filter((asset) => asset.equipmentGroup.toLowerCase() === group.toLowerCase());
    setIsSaving(true);
    try {
      const { parseAssetWorkbook } = await import("@/lib/asset-excel");
      const importedRows = parseAssetWorkbook(await file.arrayBuffer(), fieldLabels);
      const existingById = new Map(currentGroupAssets.map((asset) => [asset.id, asset]));
      const importedIds = new Set<string>();
      for (const row of importedRows) {
        if (!row.id) continue;
        if (importedIds.has(row.id)) throw new Error(`Record ID "${row.id}" appears more than once.`);
        importedIds.add(row.id);
        const existing = existingById.get(row.id);
        if (!existing) {
          throw new Error(`Record ID "${row.id}" on spreadsheet row ${row.sourceRow} does not belong to this page.`);
        }
      }

      setPendingImport({ rows: importedRows, group, groupLabel });
    } catch (error) {
      setNotice(`Excel import failed: ${errorDetails(error)}`);
    } finally {
      setIsSaving(false);
    }
  }

  async function confirmExcelImport() {
    if (!pendingImport) return;
    const { rows: importedRows, group } = pendingImport;
    const currentGroupAssets = assets.filter((asset) => asset.equipmentGroup.toLowerCase() === group.toLowerCase());
    const existingById = new Map(currentGroupAssets.map((asset) => [asset.id, asset]));
    const importedIds = new Set<string>();
    for (const row of importedRows) {
      if (!row.id) continue;
      if (importedIds.has(row.id)) {
        setNotice(`Excel import failed: Record ID "${row.id}" appears more than once.`);
        return;
      }
      importedIds.add(row.id);
      if (!existingById.has(row.id)) {
        setNotice(`Excel import failed: Record ID "${row.id}" on spreadsheet row ${row.sourceRow} does not belong to this page.`);
        return;
      }
    }
    const newRows = importedRows.filter((row) => !row.id).length;
    const updatedRows = importedRows.length - newRows;
    setIsSaving(true);
    try {
      let nextSortOrder = Math.max(0, ...currentGroupAssets.map((asset) => asset.sortOrder ?? 0));
      const nextAssets = [...assets];
      const importedAssets: Asset[] = [];
      for (const row of importedRows) {
        const existing = row.id ? existingById.get(row.id) : undefined;
        const id = existing?.id ?? `zzzz-asset-${Date.now()}-${crypto.randomUUID()}`;
        const asset: Asset = {
          ...row.fields,
          id,
          equipmentGroup: group,
          sortOrder: existing?.sortOrder ?? ++nextSortOrder,
          remarksColor: existing?.remarksColor ?? "",
        };
        const existingIndex = nextAssets.findIndex((item) => item.id === id);
        if (existingIndex >= 0) nextAssets[existingIndex] = asset;
        else nextAssets.push(asset);
        importedAssets.push(asset);
      }

      if (firebaseConfigured) {
        await saveAssetsToFirestore(importedAssets, group);
      }
      persistLocally(nextAssets);
      setAssets(nextAssets);
      setSearch("");
      setLocationFilter("All locations");
      setPhaseFilter("All phases");
      setPage(1);
      setNotice(`Imported ${importedRows.length} rows: ${updatedRows} updated, ${newRows} added.`);
      setPendingImport(null);
    } catch (error) {
      setNotice(`Excel import failed: ${errorDetails(error)}`);
    } finally {
      setIsSaving(false);
    }
  }

  async function handleSave(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!editingAsset) return;
    const isNewAsset = !assets.some((item) => item.id === editingAsset.id);
    const nextSortOrder = Math.max(0, ...assets.map((item) => item.sortOrder ?? 0)) + 1;
    const asset: Asset = {
      ...editingAsset,
      id: editingAsset.id || `zzzz-asset-${Date.now()}-${crypto.randomUUID()}`,
      sortOrder: editingAsset.sortOrder ?? nextSortOrder,
    };
    const nextAssets = assets.some((item) => item.id === asset.id)
      ? assets.map((item) => item.id === asset.id ? asset : item)
      : [...assets, asset];

    setIsSaving(true);
    try {
      const previousAsset = assets.find((item) => item.id === asset.id);
      if (firebaseConfigured) await saveAssetToFirestore(asset, previousAsset?.equipmentGroup);
      persistLocally(nextAssets);
      setAssets(nextAssets);
      if (isNewAsset) {
        setSearch("");
        setLocationFilter("All locations");
        setPhaseFilter("All phases");
        setPage(Math.floor(assets.length / pageSize) + 1);
      }
      setEditingAsset(null);
      setNotice(firebaseConfigured ? "Changes saved to Firestore." : "Changes saved on this device.");
    } catch (error) {
      const group = asset.equipmentGroup.trim().replace(/[^a-zA-Z0-9_-]+/g, "-").replace(/^-+|-+$/g, "") || "Unassigned";
      setNotice(`Save failed for Firestore collection “${group}”: ${errorDetails(error)}`);
    } finally {
      setIsSaving(false);
    }
  }

  async function handleDelete(asset: Asset) {
    if (!window.confirm(`Delete ${asset.equipmentCode || "this asset"}? This cannot be undone.`)) return;
    try {
      if (firebaseConfigured) await deleteAssetFromFirestore(asset);
      const nextAssets = assets.filter((item) => item.id !== asset.id);
      persistLocally(nextAssets);
      setAssets(nextAssets);
      setNotice("Asset removed from the register.");
    } catch (error) {
      const group = asset.equipmentGroup.trim().replace(/[^a-zA-Z0-9_-]+/g, "-").replace(/^-+|-+$/g, "") || "Unassigned";
      setNotice(`Delete failed for Firestore collection “${group}”: ${errorDetails(error)}`);
    }
  }

  async function handleRemarksColorChange(asset: Asset, color: "" | "red" | "green") {
    const updatedAsset = { ...asset, remarksColor: color };
    const nextAssets = assets.map((item) => item.id === asset.id ? updatedAsset : item);
    try {
      if (firebaseConfigured) await saveAssetToFirestore(updatedAsset);
      persistLocally(nextAssets);
      setAssets(nextAssets);
    } catch (error) {
      setNotice(`Could not save the Remarks color: ${errorDetails(error)}`);
    }
  }

  const fanCoilCount = assets.filter((asset) => asset.equipmentType === "Fan Coil Unit").length;

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#register" aria-label="MEPF asset register home">
          <span>MEPF <span className="brand-divider">/</span> FACILITIES</span>
        </a>
        <div className="topbar-right">
          <span className={`connection-state ${cloudStatus === "connected" ? "is-cloud" : "is-local"}`}><span className="state-dot" />{cloudStatus === "connected" ? "Cloud connected" : cloudStatus === "connecting" ? "Connecting..." : cloudStatus === "error" ? "Cloud sync error" : "Local workspace"}</span>
          <span className="year-chip">2022 <span>/</span> WMHC</span>
        </div>
      </header>

      <nav className="workbook-nav" aria-label="Workbook pages">
        <div className="sheet-tabs" role="tablist" aria-label="Worksheets">
          {sheets.map((sheet) => {
            const tabTone = sheet.id === "hvac-4" ? "green" : sheet.id === "ps-1" ? "yellow" : ["hvac-1", "hvac-2", "hvac-3"].includes(sheet.id) ? "red" : sheet.kind === "group" ? "custom" : "neutral";
            return <div className="sheet-tab-wrap" key={sheet.id} role="presentation">
              <button
                className={`sheet-tab sheet-tab-${tabTone} ${activeSheetId === sheet.id ? "is-active" : ""}`}
                id={`tab-${sheet.id}`}
                role="tab"
                aria-selected={activeSheetId === sheet.id}
                aria-controls="workbook-page"
                disabled={isSaving}
                onClick={() => { setActiveSheetId(sheet.id); setPage(1); setSearch(""); setLocationFilter("All locations"); setPhaseFilter("All phases"); }}
              >{sheet.label}</button>
            </div>;
          })}
        </div>
        <button className="add-sheet-button" onClick={() => setAddingPage(true)} disabled={isSaving} aria-label="Add page" title="Add page"><Plus size={16} /></button>
      </nav>

      <section className="page-content" id="workbook-page" role="tabpanel" aria-labelledby={`tab-${activeSheet.id}`}>
        <div className="page-heading">
          <div>
            <h1>{activeSheet.kind === "index" ? "INDEX" : activeSheet.label}</h1>
            <p className="page-subtitle">{activeSheet.kind === "terms" ? "Reference guide for MEPF equipment terms and condition ratings" : activeSheet.kind === "life" ? "Service life and replacement planning by equipment" : activeSheet.kind === "group" ? `Equipment records for ${activeSheet.group ?? activeSheet.label}` : "Equipment groups and descriptions"}</p>
          </div>
          <div className="page-heading-actions">
            {!defaultSheets.some((sheet) => sheet.id === activeSheet.id) && <button className="delete-page-button" onClick={() => void deleteWorkbookPage(activeSheet)} disabled={isSaving}><Trash2 size={15} /> Delete page</button>}
            {activeSheet.kind === "group" && <>
              <button className="icon-button excel-template-button" onClick={() => void downloadExcelTemplate()} disabled={isSaving || (firebaseConfigured && cloudStatus !== "connected")} title="Download Excel template"><ArrowDownToLine size={16} /> Template</button>
              <button className="icon-button excel-upload-button" onClick={() => importFileInput.current?.click()} disabled={isSaving || (firebaseConfigured && cloudStatus !== "connected")} title="Upload completed Excel workbook"><Upload size={16} /> Upload Excel</button>
              <input
                ref={importFileInput}
                className="sr-only"
                type="file"
                accept=".xlsx,.xls"
                aria-label="Choose an Excel workbook to import"
                onChange={(event) => void handleExcelUpload(event)}
              />
            </>}
            {activeSheet.kind === "index" && <button className="primary-button" onClick={() => { setReferenceFormError(""); setEditingIndexRow({ id: "", number: "", equipmentGroup: "", description: "" }); }} disabled={isSaving}><Plus size={17} /> Add index row</button>}
            {activeSheet.kind === "terms" && <button className="primary-button" onClick={() => { setReferenceFormError(""); setEditingDefinition({ term: "", definition: "" }); setEditingDefinitionOriginalTerm(null); }} disabled={isSaving}><Plus size={17} /> Add term</button>}
            {(activeSheet.kind === "group" || activeSheet.kind === "life") && <button className="primary-button" onClick={startAddingAsset} disabled={isSaving}><Plus size={17} /> Add equipment</button>}
          </div>
        </div>

        {notice && <div className="notice page-notice" role="status"><Check size={14} />{notice}<button onClick={() => setNotice("")} aria-label="Dismiss message"><X size={14} /></button></div>}

        {activeSheet.kind === "terms" ? (
          <section className="reference-sheet" aria-label="Definitions">
            <table className="reference-table">
              <thead><tr><th>TERM</th><th>DEFINITION</th><th className="reference-action-heading">ACTIONS</th></tr></thead>
              <tbody>{definitionRows.map((item) => <tr key={item.term}><th scope="row">{item.term}</th><td>{item.definition}</td><td className="reference-actions"><button className="row-icon-button" onClick={() => { setReferenceFormError(""); setEditingDefinition({ ...item }); setEditingDefinitionOriginalTerm(item.term); }} disabled={isSaving} aria-label={`Edit ${item.term}`} title="Edit definition"><Pencil size={14} /></button><button className="row-icon-button delete-action" onClick={() => void deleteDefinition(item)} disabled={isSaving} aria-label={`Delete ${item.term}`} title="Delete definition"><Trash2 size={14} /></button></td></tr>)}</tbody>
            </table>
          </section>
        ) : activeSheet.kind === "life" ? (
          <section className="reference-sheet" aria-label="MEPF life expectancy">
            <table className="reference-table life-table">
              <thead><tr><th>EQUIPMENT CODE</th><th>SERVICED AREA</th><th>INSTALLATION YEAR</th><th>NORMAL LIFE</th><th>PRESENT CONDITION</th><th>ADJUSTED LIFE</th><th>REPLACEMENT YEAR</th><th className="reference-action-heading">ACTIONS</th></tr></thead>
              <tbody>{filteredAssets.map((asset) => <tr key={asset.id}><td>{asset.equipmentCode || "—"}</td><td>{asset.servicedArea || "—"}</td><td>{asset.installationYear || "—"}</td><td>{asset.normalLife || "—"}</td><td>{asset.presentCondition || "—"}</td><td>{asset.adjustedLife || "—"}</td><td>{asset.replacementYear || "—"}</td><td className="reference-actions"><button className="row-icon-button" onClick={() => setEditingAsset({ ...asset })} disabled={isSaving} aria-label={`Edit ${asset.equipmentCode || asset.servicedArea}`} title="Edit equipment"><Pencil size={14} /></button><button className="row-icon-button delete-action" onClick={() => void handleDelete(asset)} disabled={isSaving} aria-label={`Delete ${asset.equipmentCode || asset.servicedArea}`} title="Delete equipment"><Trash2 size={14} /></button></td></tr>)}</tbody>
            </table>
            {filteredAssets.length === 0 && <p className="reference-empty">No assets are available for life expectancy planning.</p>}
          </section>
        ) : activeSheet.kind === "index" ? (
          <section className="reference-sheet" aria-label="Equipment group index">
            <table className="reference-table index-table">
              <thead><tr><th>NO.</th><th>EQPT GROUP</th><th>DESCRIPTION</th><th className="reference-action-heading">ACTIONS</th></tr></thead>
              <tbody>{indexRows.map((row) => <tr key={row.id}><td className="index-number">{row.number}</td><td className="index-group">{row.equipmentGroup}</td><td>{row.description}</td><td className="reference-actions"><button className="row-icon-button" onClick={() => { setReferenceFormError(""); setEditingIndexRow({ ...row }); }} disabled={isSaving} aria-label={`Edit ${row.description}`} title="Edit index row"><Pencil size={14} /></button><button className="row-icon-button delete-action" onClick={() => void deleteIndexRow(row)} disabled={isSaving} aria-label={`Delete ${row.description}`} title="Delete index row"><Trash2 size={14} /></button></td></tr>)}</tbody>
            </table>
          </section>
        ) : <>
        <div className="summary-row" aria-label="Asset summary">
          <div className="summary-item"><span className="summary-icon summary-icon-green"><ClipboardList size={17} /></span><span className="summary-label">REGISTERED ASSETS</span><strong>{assets.length.toLocaleString()}</strong></div>
          <div className="summary-item"><span className="summary-icon summary-icon-orange"><Fan size={17} strokeWidth={1.8} /></span><span className="summary-label">FAN COIL UNITS</span><strong>{fanCoilCount.toLocaleString()}</strong></div>
          <div className="summary-item"><span className="summary-icon summary-icon-blue"><MapPin size={17} strokeWidth={1.8} /></span><span className="summary-label">LOCATIONS COVERED</span><strong>{locations.length.toLocaleString()}</strong></div>
          <div className="summary-footnote">Last updated <strong>just now</strong></div>
        </div>

        <section className="register-panel" aria-label="Equipment register">
          <div className="table-toolbar">
            <div className="table-title-group"><h2>Equipment</h2><span className="record-count">{filteredAssets.length} records</span></div>
            <div className="table-actions">
              <label className="search-field"><Search size={16} /><input aria-label="Search equipment" placeholder="Search equipment..." value={search} onChange={(event) => { setSearch(event.target.value); setPage(1); }} />{search && <button className="clear-search" onClick={() => { setSearch(""); setPage(1); }} aria-label="Clear search"><X size={14} /></button>}</label>
              <label className="select-field"><select aria-label="Filter by location" value={locationFilter} onChange={(event) => { setLocationFilter(event.target.value); setPage(1); }}><option>All locations</option>{locations.map((location) => <option key={location}>{location}</option>)}</select><ChevronDown size={14} /></label>
              <label className="select-field phase-select"><select aria-label="Filter by phase" value={phaseFilter} onChange={(event) => { setPhaseFilter(event.target.value); setPage(1); }}><option>All phases</option>{phases.map((phase) => <option key={phase}>{phase}</option>)}</select><ChevronDown size={14} /></label>
              <button className="icon-button export-button" onClick={exportCsv} title="Export visible rows as CSV" aria-label="Export CSV"><ArrowDownToLine size={16} /><span>Export</span></button>
            </div>
          </div>

          <div className="table-scroll">
            <table className="asset-table">
              <thead>
                <tr className="group-row">
                  <th rowSpan={2}>EQPT GROUP</th><th rowSpan={2}>EQPT TYPE</th><th rowSpan={2}>EQPT TAG</th><th rowSpan={2}>EQPT LOCATION</th><th rowSpan={2}>SERVICED AREA</th><th rowSpan={2}>SEQUENCE NO.</th><th rowSpan={2}>EQPT CODE</th><th rowSpan={2}>BRAND</th><th rowSpan={2}>CAPACITY</th><th rowSpan={2}>HOTEL ASSET CODE</th><th rowSpan={2}>INSTALLATION YEAR</th><th colSpan={3} className="life-group">LIFE EXPECTANCY</th><th rowSpan={2}>REPLACEMENT YEAR</th><th rowSpan={2}>ZONING</th><th rowSpan={2}>PHASING</th><th rowSpan={2} className="remarks-heading">REMARKS</th><th rowSpan={2} className="action-heading"><span className="sr-only">Actions</span></th>
                </tr>
                <tr className="subhead-row"><th>NORMAL</th><th>PRESENT CONDITION</th><th>ADJUSTED</th></tr>
              </thead>
              <tbody>
                {visibleAssets.map((asset, index) => (
                  <tr key={asset.id}>
                    <td><span className="group-tag">{asset.equipmentGroup || "—"}</span></td><td>{asset.equipmentType || "—"}</td><td className="center-cell">{asset.equipmentTag || "—"}</td><td>{asset.equipmentLocation || "—"}</td><td>{asset.servicedArea || "—"}</td><td className="center-cell">{asset.sequenceNumber || "—"}</td><td className="code-cell">{asset.equipmentCode || "—"}</td><td>{asset.brand || "—"}</td><td className="center-cell">{asset.capacity || "—"}</td><td className="center-cell">{asset.hotelAssetCode || "—"}</td><td className="center-cell">{asset.installationYear || "—"}</td><td className="center-cell">{asset.normalLife || "—"}</td><td className="center-cell">{asset.presentCondition || "—"}</td><td className="center-cell">{asset.adjustedLife || "—"}</td><td className="center-cell">{asset.replacementYear || "—"}</td><td className="center-cell"><span className="zone-tag">{asset.zoning || "—"}</span></td><td className="center-cell">{asset.phasing || "—"}</td><td className={`remarks-cell remarks-${asset.remarksColor || "none"}`}><span>{asset.remarks}</span><select aria-label={`Remarks color for ${asset.equipmentCode || asset.servicedArea}`} value={asset.remarksColor ?? ""} onChange={(event) => void handleRemarksColorChange(asset, event.target.value as "" | "red" | "green")}><option value="">No color</option><option value="red">Red</option><option value="green">Green</option></select></td>
                    <td className="row-actions"><button className="row-icon-button" onClick={() => setEditingAsset({ ...asset })} aria-label={`Edit ${asset.equipmentCode || `asset ${index + 1}`}`} title="Edit asset"><Pencil size={14} /></button><button className="row-icon-button delete-action" onClick={() => void handleDelete(asset)} aria-label={`Delete ${asset.equipmentCode || `asset ${index + 1}`}`} title="Delete asset"><Trash2 size={14} /></button></td>
                  </tr>
                ))}
                {visibleAssets.length === 0 && <tr><td className="empty-state" colSpan={19}>No equipment matches these filters.</td></tr>}
              </tbody>
            </table>
          </div>
          <footer className="table-footer">
            <span>Showing <strong>{firstVisibleRow}–{lastVisibleRow}</strong> of <strong>{filteredAssets.length}</strong> assets</span>
            <div className="pagination-controls" aria-label="Table pagination">
              <button className="page-button" onClick={() => setPage(currentPage - 1)} disabled={currentPage <= 1} aria-label="Previous page"><ChevronLeft size={15} /></button>
              <span>Page <strong>{currentPage}</strong> of <strong>{pageCount}</strong></span>
              <button className="page-button" onClick={() => setPage(currentPage + 1)} disabled={currentPage >= pageCount} aria-label="Next page"><ChevronRight size={15} /></button>
            </div>
            <span className="footer-note"><span className="state-dot" />{cloudStatus === "connected" ? "Live Firestore sync" : cloudStatus === "error" ? "Firestore unavailable" : "Saved in this browser"}</span>
          </footer>
        </section>
        </>}
        <div className="page-footer">MEPF ASSET MANAGEMENT <span>•</span> WMHC ENGINEERING</div>
      </section>

      {addingPage && <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) setAddingPage(false); }}>
        <section className="asset-modal add-page-modal" role="dialog" aria-modal="true" aria-labelledby="add-page-title">
          <header className="modal-header"><div><span className="modal-kicker">WORKBOOK</span><h2 id="add-page-title">Add page</h2></div><button className="icon-button modal-close" onClick={() => setAddingPage(false)} aria-label="Close dialog"><X size={18} /></button></header>
          <form onSubmit={(event) => void addWorkbookPage(event)}>
            <div className="add-page-form"><label className="form-field"><span>Page name</span><input autoFocus value={newPageName} onChange={(event) => setNewPageName(event.target.value)} placeholder="e.g. Plumbing-1" required /></label><p>Equipment added on this page will use the page name as its equipment group.</p></div>
            <footer className="modal-footer"><button type="button" className="secondary-button" onClick={() => setAddingPage(false)} disabled={isSaving}>Cancel</button><button type="submit" className="primary-button" disabled={isSaving}>{isSaving ? "Saving..." : <><Plus size={16} /> Create page</>}</button></footer>
          </form>
        </section>
      </div>}

      {editingIndexRow && <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget && !isSaving) setEditingIndexRow(null); }}>
        <section className="asset-modal reference-modal" role="dialog" aria-modal="true" aria-labelledby="index-row-title">
          <header className="modal-header"><div><span className="modal-kicker">INDEX</span><h2 id="index-row-title">{editingIndexRow.id ? "Edit index row" : "Add index row"}</h2></div><button className="icon-button modal-close" onClick={() => setEditingIndexRow(null)} disabled={isSaving} aria-label="Close dialog"><X size={18} /></button></header>
          <form onSubmit={(event) => void handleIndexRowSave(event)}>
            <div className="form-grid reference-form-grid">
              <label className="form-field"><span>No.</span><input autoFocus value={editingIndexRow.number} onChange={(event) => { setReferenceFormError(""); setEditingIndexRow({ ...editingIndexRow, number: event.target.value }); }} placeholder="e.g. 1" inputMode="numeric" /></label>
              <label className="form-field"><span>Equipment group</span><input value={editingIndexRow.equipmentGroup} onChange={(event) => { setReferenceFormError(""); setEditingIndexRow({ ...editingIndexRow, equipmentGroup: event.target.value }); }} placeholder="e.g. HVAC-1" /></label>
              <label className="form-field field-full"><span>Description</span><textarea value={editingIndexRow.description} onChange={(event) => { setReferenceFormError(""); setEditingIndexRow({ ...editingIndexRow, description: event.target.value }); }} placeholder="Enter equipment group or equipment description" required /></label>
              {referenceFormError && <p className="reference-form-error field-full" role="alert">{referenceFormError}</p>}
            </div>
            <footer className="modal-footer"><button type="button" className="secondary-button" onClick={() => setEditingIndexRow(null)} disabled={isSaving}>Cancel</button><button type="submit" className="primary-button" disabled={isSaving}>{isSaving ? "Saving..." : <><Check size={16} /> Save index row</>}</button></footer>
          </form>
        </section>
      </div>}

      {editingDefinition && <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget && !isSaving) setEditingDefinition(null); }}>
        <section className="asset-modal reference-modal" role="dialog" aria-modal="true" aria-labelledby="definition-title">
          <header className="modal-header"><div><span className="modal-kicker">DEFINITION OF TERMS</span><h2 id="definition-title">{editingDefinitionOriginalTerm !== null ? "Edit definition" : "Add definition"}</h2></div><button className="icon-button modal-close" onClick={() => { setEditingDefinition(null); setEditingDefinitionOriginalTerm(null); }} disabled={isSaving} aria-label="Close dialog"><X size={18} /></button></header>
          <form onSubmit={(event) => void handleDefinitionSave(event)}>
            <div className="form-grid reference-form-grid">
              <label className="form-field field-full"><span>Term</span><input autoFocus value={editingDefinition.term} onChange={(event) => { setReferenceFormError(""); setEditingDefinition({ ...editingDefinition, term: event.target.value }); }} placeholder="Enter term" required /></label>
              <label className="form-field field-full"><span>Definition</span><textarea value={editingDefinition.definition} onChange={(event) => { setReferenceFormError(""); setEditingDefinition({ ...editingDefinition, definition: event.target.value }); }} placeholder="Enter definition" required /></label>
              {referenceFormError && <p className="reference-form-error field-full" role="alert">{referenceFormError}</p>}
            </div>
            <footer className="modal-footer"><button type="button" className="secondary-button" onClick={() => { setEditingDefinition(null); setEditingDefinitionOriginalTerm(null); }} disabled={isSaving}>Cancel</button><button type="submit" className="primary-button" disabled={isSaving}>{isSaving ? "Saving..." : <><Check size={16} /> Save definition</>}</button></footer>
          </form>
        </section>
      </div>}

      {editingAsset && <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget && !isSaving) setEditingAsset(null); }}>
        <section className="asset-modal" role="dialog" aria-modal="true" aria-labelledby="modal-title">
          <header className="modal-header"><div><span className="modal-kicker">ASSET REGISTER</span><h2 id="modal-title">{editingAsset.id ? "Edit equipment" : "Add equipment"}</h2></div><button className="icon-button modal-close" onClick={() => setEditingAsset(null)} aria-label="Close dialog"><X size={18} /></button></header>
          <form onSubmit={(event) => void handleSave(event)}>
            <div className="form-grid">{assetFields.map((field) => <label className={`form-field ${field === "equipmentCode" || field === "servicedArea" ? "field-wide" : ""}`} key={field}><span>{fieldLabels[field]}</span><input autoFocus={field === "equipmentGroup" && !editingAsset.id} value={editingAsset[field]} onChange={(event) => setEditingAsset({ ...editingAsset, [field]: event.target.value })} placeholder={field === "equipmentCode" ? "e.g. HVAC-4-FCU-BMT-Guest Entrance-1" : "Enter value"} /></label>)}</div>
            <footer className="modal-footer"><button type="button" className="secondary-button" onClick={() => setEditingAsset(null)} disabled={isSaving}>Cancel</button><button type="submit" className="primary-button" disabled={isSaving}>{isSaving ? "Saving..." : <><Check size={16} /> Save equipment</>}</button></footer>
          </form>
        </section>
      </div>}

      {pendingImport && <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget && !isSaving) setPendingImport(null); }}>
        <section className="asset-modal import-preview-modal" role="dialog" aria-modal="true" aria-labelledby="import-preview-title">
          <header className="modal-header"><div><span className="modal-kicker">CHECK BEFORE SAVING</span><h2 id="import-preview-title">Preview Excel import · {pendingImport.groupLabel}</h2></div><button className="icon-button modal-close" onClick={() => setPendingImport(null)} disabled={isSaving} aria-label="Close preview"><X size={18} /></button></header>
          <div className="import-preview-content">
            <p className="import-preview-note">Review each spreadsheet row and its detected column values. Equipment group will be set to <strong>{pendingImport.groupLabel}</strong> automatically. Blank cells stay blank; nothing is saved until you confirm.</p>
            <div className="import-preview-scroll">
              <table className="import-preview-table">
                <thead><tr><th>Spreadsheet row</th><th>Record ID</th>{assetFields.map((field) => <th key={field}>{fieldLabels[field]}</th>)}</tr></thead>
                <tbody>{pendingImport.rows.map((row, index) => <tr key={`${row.id}-${index}`}>
                  <td>{row.sourceRow}</td>
                  <td>{row.id || "New"}</td>
                  {assetFields.map((field) => <td key={field} className={!row.fields[field] ? "import-preview-empty" : ""}>
                    {field === "equipmentGroup" ? pendingImport.group : row.fields[field] || "—"}
                  </td>)}
                </tr>)}</tbody>
              </table>
            </div>
            <p className="import-preview-summary">{pendingImport.rows.length} rows · {pendingImport.rows.filter((row) => row.id).length} updates · {pendingImport.rows.filter((row) => !row.id).length} new records. Other records will not be deleted.</p>
          </div>
          <footer className="modal-footer"><button type="button" className="secondary-button" onClick={() => setPendingImport(null)} disabled={isSaving}>Cancel</button><button type="button" className="primary-button" onClick={() => void confirmExcelImport()} disabled={isSaving}>{isSaving ? "Importing..." : <><Check size={16} /> Confirm import</>}</button></footer>
        </section>
      </div>}
    </main>
  );
}
