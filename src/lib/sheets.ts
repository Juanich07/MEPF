export type WorkbookSheet = {
  id: string;
  label: string;
  kind: "index" | "life" | "terms" | "group";
  group?: string;
};

export const defaultSheets: WorkbookSheet[] = [
  { id: "index", label: "Index", kind: "index" },
  { id: "life-expectancy", label: "MEPF Life Expectancy", kind: "life" },
  { id: "definitions", label: "Definition of Terms", kind: "terms" },
  { id: "hvac-1", label: "1 HVAC-1", kind: "group", group: "HVAC-1" },
  { id: "hvac-2", label: "2 HVAC-2", kind: "group", group: "HVAC-2" },
  { id: "hvac-3", label: "3 HVAC-3", kind: "group", group: "HVAC-3" },
  { id: "hvac-4", label: "4 HVAC-4", kind: "group", group: "HVAC-4" },
  { id: "ps-1", label: "7 PS-1", kind: "group", group: "PS-1" },
];

export const termDefinitions = [
  { term: "Equipment group", definition: "Classification or system group the equipment belongs to." },
  { term: "Equipment type", definition: "The kind of equipment, such as a fan coil unit or pump." },
  { term: "Equipment tag", definition: "Short identifier used to recognize the equipment type." },
  { term: "Equipment location", definition: "Building or site where the equipment is installed." },
  { term: "Serviced area", definition: "Room, zone, or area served by the equipment." },
  { term: "Equipment sequence number", definition: "Number distinguishing similar equipment in a group or area." },
  { term: "Equipment code", definition: "Combined unique code used to identify the equipment record." },
  { term: "Capacity", definition: "Rated output or size of the equipment, with its unit." },
  { term: "Installation year", definition: "Year the equipment was installed or commissioned." },
  { term: "Normal life expectancy", definition: "Expected service life under normal operating conditions." },
  { term: "Present condition", definition: "Observed condition used when assessing remaining service life." },
  { term: "Adjusted life expectancy", definition: "Expected life adjusted for the recorded condition and use." },
  { term: "Replacement year", definition: "Planned year for replacement based on life and condition." },
  { term: "Zoning", definition: "Facility zone or building reference associated with the asset." },
  { term: "Phasing", definition: "Planned project phase for work on the asset." },
  { term: "Remarks", definition: "Additional notes or procurement/status information." },
];