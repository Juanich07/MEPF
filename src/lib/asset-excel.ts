import * as XLSX from "@e965/xlsx";
import { assetFields, type Asset, type AssetField } from "@/lib/assets";

export type ImportedAssetRow = {
  id: string;
  fields: Record<AssetField, string>;
  sourceRow: number;
};

const excelFieldLabels: Record<AssetField, string> = {
  equipmentGroup: "EQPT GROUP",
  equipmentType: "EQPT TYPE",
  equipmentTag: "EQPT TAG",
  equipmentLocation: "EQPT LOCATION",
  servicedArea: "SERVICED AREA",
  sequenceNumber: "SEQUENCE NO.",
  equipmentCode: "EQPT CODE",
  brand: "BRAND",
  capacity: "CAPACITY",
  hotelAssetCode: "HOTEL ASSET CODE",
  installationYear: "INSTALLATION YEAR",
  normalLife: "NORMAL",
  presentCondition: "PRESENT CONDITION",
  adjustedLife: "ADJUSTED",
  replacementYear: "REPLACEMENT YEAR",
  zoning: "ZONING",
  phasing: "PHASING",
  remarks: "REMARKS",
};

function cellText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value).trim();
}

export async function downloadAssetTemplate(
  group: string,
  assets: Asset[],
): Promise<void> {
  const fileGroup = group.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "equipment";
  const { default: writeXlsxFile } = await import("write-excel-file/browser");
  const border = { borderColor: "#D9E1DB", borderStyle: "thin" as const };
  const headerStyle = {
    ...border,
    backgroundColor: "#EDF2EE",
    fontWeight: "bold" as const,
    textColor: "#44534A",
    wrap: true,
    align: "center" as const,
    alignVertical: "center" as const,
  };
  const lifeHeaderStyle = {
    ...border,
    backgroundColor: "#E6EEE8",
    fontWeight: "bold" as const,
    textColor: "#44534A",
    align: "center" as const,
    wrap: true,
    alignVertical: "center" as const,
  };
  const bodyStyle = {
    ...border,
    alignVertical: "center" as const,
  };
  const recordIdStyle = {
    ...bodyStyle,
    backgroundColor: "#F3F6F3",
    textColor: "#708078",
  };
  const blankRows = Array.from({ length: 20 }, () =>
    Array.from({ length: assetFields.length + 1 }, () => ({ value: "", ...bodyStyle })),
  );
  const assetsData = [
    [
      { value: "Record ID", rowSpan: 2, ...headerStyle },
      ...assetFields.slice(0, 11).map((field) => ({
        value: excelFieldLabels[field],
        rowSpan: 2,
        ...headerStyle,
      })),
      { value: "LIFE EXPECTANCY", columnSpan: 3, ...lifeHeaderStyle },
      null,
      null,
      ...assetFields.slice(14).map((field) => ({
        value: excelFieldLabels[field],
        rowSpan: 2,
        ...headerStyle,
      })),
    ],
    [
      ...Array.from({ length: 12 }, () => null),
      ...assetFields.slice(11, 14).map((field) => ({
        value: excelFieldLabels[field],
        ...headerStyle,
      })),
      null,
      null,
      null,
      null,
    ],
    ...assets.map((asset) => [
      { value: asset.id, ...recordIdStyle },
      ...assetFields.map((field) => ({ value: asset[field], ...bodyStyle })),
    ]),
    ...blankRows,
  ];

  await writeXlsxFile([
    {
      sheet: "Assets",
      orientation: "landscape",
      showGridLines: false,
      stickyRowsCount: 2,
      columns: [
        { width: 1 },
        { width: 17 },
        { width: 21 },
        { width: 14 },
        { width: 21 },
        { width: 15 },
        { width: 15 },
        { width: 28 },
        { width: 17 },
        { width: 17 },
        { width: 15 },
        { width: 21 },
        { width: 17 },
        { width: 21 },
        { width: 20 },
        { width: 13 },
        { width: 16 },
        { width: 24 },
      ],
      data: assetsData,
    },
    {
      sheet: "Instructions",
      columns: [{ width: 100 }],
      stickyRowsCount: 1,
      data: [
        [{ value: "MEPF equipment import", ...lifeHeaderStyle }],
        ["The Assets sheet uses the same column layout as the equipment register."],
        ["Fill in the blank rows to add equipment, or edit existing rows to update them."],
        ["The template includes a hidden Record ID to match existing rows. Leave it unchanged; new blank rows are added as equipment."],
        ["Blank rows are ignored. Uploading does not delete records that are missing from the workbook."],
        ["Equipment group is assigned from the page you upload into."],
      ],
    },
  ], {
    fontFamily: "Arial",
    fontSize: 10,
    features: [{
      files: {
        transform: {
          "xl/worksheets/sheet{id}.xml": {
            transformElementAttributes(tagName, attributes, index, _options, properties) {
              if (properties.sheetIndex === 0 && tagName === "col" && index === 0) {
                return { ...attributes, hidden: 1 };
              }
              return attributes;
            },
          },
        },
      },
    }],
  }).toFile(`mepf-${fileGroup}-template.xlsx`);
}

export function parseAssetWorkbook(
  data: ArrayBuffer,
  fieldLabels: Record<AssetField, string>,
): ImportedAssetRow[] {
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  const sheet = workbook.Sheets.Assets;
  if (!sheet) throw new Error('The workbook must contain an "Assets" sheet.');

  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: "",
    raw: false,
  });
  const requiredHeaders = ["Record ID"];
  const headerRowIndex = rows.findIndex((row, index) => {
    const headers = row.map(cellText);
    const nextRow = rows[index + 1]?.map(cellText) ?? [];
    const combinedHeaders = headers.map((header, columnIndex) =>
      !header || header.toLowerCase() === "life expectancy"
        ? nextRow[columnIndex] || header
        : header,
    );
    return assetFields.every((field) => {
      const aliases = [excelFieldLabels[field], fieldLabels[field]].map((header) => header.toLowerCase());
      return combinedHeaders.some((header) => aliases.includes(header.toLowerCase()));
    });
  });
  if (headerRowIndex < 0) throw new Error("The Assets sheet is missing its column headers.");

  const rowHeaders = rows[headerRowIndex].map(cellText);
  const nextRowHeaders = rows[headerRowIndex + 1]?.map(cellText) ?? [];
  const headers = rowHeaders.map((header, index) =>
    !header || header.toLowerCase() === "life expectancy"
      ? nextRowHeaders[index] || header
      : header,
  );
  if (!headers.some((header) => header.toLowerCase() === "record id")) {
    const mergedIdHeader = rows[headerRowIndex - 1]?.[0];
    if (cellText(mergedIdHeader).toLowerCase() === "record id") headers[0] = "Record ID";
  }
  const headerIndexes = new Map<string, number>();
  headers.forEach((header, index) => {
    const normalized = header.toLowerCase();
    if (normalized && headerIndexes.has(normalized)) {
      throw new Error(`The Assets sheet has a duplicate column: ${header}.`);
    }
    if (normalized) headerIndexes.set(normalized, index);
  });

  const missingHeaders = requiredHeaders.filter((header) =>
    !headerIndexes.has(header.toLowerCase()),
  );
  if (missingHeaders.length > 0) {
    throw new Error(`Missing required columns: ${missingHeaders.join(", ")}.`);
  }

  function columnIndex(header: string): number {
    const index = headerIndexes.get(header.toLowerCase());
    if (index === undefined) throw new Error(`Missing required column: ${header}.`);
    return index;
  }

  const rowsByField = assetFields.map((field) => {
    const index = headerIndexes.get(excelFieldLabels[field].toLowerCase())
      ?? headerIndexes.get(fieldLabels[field].toLowerCase());
    if (index === undefined) throw new Error(`Missing required column: ${excelFieldLabels[field]}.`);
    return { field, index };
  });
  const idColumn = columnIndex("Record ID");
  const hasSubheaderRow = assetFields.slice(11, 14).every((field) =>
    nextRowHeaders.some((header) => header.toLowerCase() === excelFieldLabels[field].toLowerCase()),
  );
  const dataStartIndex = headerRowIndex + (hasSubheaderRow ? 2 : 1);
  const importedRows = rows.slice(dataStartIndex)
    .filter((row) => row.some((value) => cellText(value) !== ""))
    .map((row, rowIndex) => {
      const id = cellText(row[idColumn]);
      if (id.includes("/")) {
        throw new Error(`Record ID on spreadsheet row ${rowIndex + dataStartIndex + 1} cannot contain "/".`);
      }
      const fields = Object.fromEntries(rowsByField.map(({ field, index }) => [
        field,
        cellText(row[index]),
      ])) as Record<AssetField, string>;
      return { id, fields, sourceRow: rowIndex + dataStartIndex + 1 };
    });

  if (importedRows.length === 0) {
    throw new Error("The Assets sheet has no equipment rows to import.");
  }

  return importedRows;
}
