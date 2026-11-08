import * as XLSX from "@e965/xlsx";
import { assetColumnLabels, type Asset, type AssetField } from "@/lib/assets";
import type { WorkbookColumn } from "@/lib/sheets";
import type { Row } from "write-excel-file/browser";

export type ImportedAssetRow = {
  id: string;
  fields: Partial<Record<AssetField, string>>;
  customFields: Record<string, string>;
  sourceRow: number;
};

function cellText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value).trim();
}

function getColumnValue(asset: Asset, column: WorkbookColumn): string {
  return "field" in column
    ? asset[column.field]
    : asset.customFields?.[column.customFieldId] ?? "";
}

function getColumnGroups(columns: WorkbookColumn[]) {
  const lifeFields: AssetField[] = ["normalLife", "presentCondition", "adjustedLife"];
  const groups: { columns: WorkbookColumn[]; isLife: boolean }[] = [];
  for (const column of columns) {
    const isLife = "field" in column && lifeFields.includes(column.field);
    const lastGroup = groups.at(-1);
    if (isLife && lastGroup?.isLife) lastGroup.columns.push(column);
    else groups.push({ columns: [column], isLife });
  }
  return groups;
}

export async function downloadAssetTemplate(
  group: string,
  assets: Asset[],
  columns: WorkbookColumn[],
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
  const columnGroups = getColumnGroups(columns);
  const hasLifeGroup = columnGroups.some((columnGroup) => columnGroup.isLife);
  const blankRows: Row[] = Array.from({ length: 20 }, () =>
    Array.from({ length: columns.length + 1 }, (): Row[number] => ({ value: "", ...bodyStyle })),
  );
  const headerRow: Row = [
    { value: "Record ID", ...(hasLifeGroup ? { rowSpan: 2 } : {}), ...headerStyle },
  ];
  for (const columnGroup of columnGroups) {
    if (columnGroup.isLife) {
      headerRow.push(
        { value: "LIFE EXPECTANCY", columnSpan: columnGroup.columns.length, ...lifeHeaderStyle },
        ...Array.from({ length: columnGroup.columns.length - 1 }, () => null),
      );
    } else {
      headerRow.push({
        value: columnGroup.columns[0].label,
        ...(hasLifeGroup ? { rowSpan: 2 } : {}),
        ...headerStyle,
      });
    }
  }
  const assetsData: Row[] = [headerRow];
  if (hasLifeGroup) {
    const subheaderRow: Row = [null];
    for (const columnGroup of columnGroups) {
      if (columnGroup.isLife) {
        subheaderRow.push(...columnGroup.columns.map((column) => ({ value: column.label, ...headerStyle })));
      } else {
        subheaderRow.push(null);
      }
    }
    assetsData.push(subheaderRow);
  }
  assetsData.push(...assets.map((asset) => [
    { value: asset.id, ...recordIdStyle },
    ...columns.map((column) => ({ value: getColumnValue(asset, column), ...bodyStyle })),
  ]));
  assetsData.push(...blankRows);

  await writeXlsxFile([
    {
      sheet: "Assets",
      orientation: "landscape",
      showGridLines: false,
      stickyRowsCount: hasLifeGroup ? 2 : 1,
      columns: [
        { width: 1 },
        ...columns.map((column) => ({ width: Math.min(32, Math.max(12, column.label.length + 3)) })),
      ],
      data: assetsData,
    },
    {
      sheet: "Instructions",
      columns: [{ width: 100 }],
      stickyRowsCount: 1,
      data: [
        [{ value: "MEPF equipment import", ...lifeHeaderStyle }],
        ["The Assets sheet uses the same configured columns as the equipment register."],
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
  columns: WorkbookColumn[],
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
    return columns.every((column) => {
      const aliases = [
        column.label,
        ...("field" in column
          ? [assetColumnLabels[column.field], fieldLabels[column.field]]
          : []),
      ].map((header) => header.toLowerCase());
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

  const rowsByColumn = columns.map((column) => {
    const aliases = [
      column.label,
      ...("field" in column
        ? [assetColumnLabels[column.field], fieldLabels[column.field]]
        : []),
    ];
    const index = aliases
      .map((header) => headerIndexes.get(header.toLowerCase()))
      .find((headerIndex) => headerIndex !== undefined);
    if (index === undefined) throw new Error(`Missing required column: ${column.label}.`);
    return { column, index };
  });
  const idColumn = columnIndex("Record ID");
  const lifeFields: AssetField[] = ["normalLife", "presentCondition", "adjustedLife"];
  const hasSubheaderRow = columns.some((column) =>
    "field" in column && lifeFields.includes(column.field)
      && nextRowHeaders.some((header) => header.toLowerCase() === column.label.toLowerCase()),
  );
  const dataStartIndex = headerRowIndex + (hasSubheaderRow ? 2 : 1);
  const importedRows = rows.slice(dataStartIndex)
    .filter((row) => row.some((value) => cellText(value) !== ""))
    .map((row, rowIndex) => {
      const id = cellText(row[idColumn]);
      if (id.includes("/")) {
        throw new Error(`Record ID on spreadsheet row ${rowIndex + dataStartIndex + 1} cannot contain "/".`);
      }
      const fields: Partial<Record<AssetField, string>> = {};
      const customFields: Record<string, string> = {};
      for (const { column, index } of rowsByColumn) {
        if ("field" in column) fields[column.field] = cellText(row[index]);
        else customFields[column.customFieldId] = cellText(row[index]);
      }
      return { id, fields, customFields, sourceRow: rowIndex + dataStartIndex + 1 };
    });

  if (importedRows.length === 0) {
    throw new Error("The Assets sheet has no equipment rows to import.");
  }

  return importedRows;
}
