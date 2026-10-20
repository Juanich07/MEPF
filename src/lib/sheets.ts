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

export type TermDefinition = {
  term: string;
  definition: string;
};

export const termDefinitions = [
  { term: "ADJUSTED LIFE EXPECTANCY", definition: "Adjusted life expectancy based on the present condition of the equipment as assessed by the property." },
  { term: "Good", definition: "Equipment performance is within standards. Life expectancy will extend by 10%." },
  { term: "Fair", definition: "Equipment is not performing well but is manageable and safe to use. Life expectancy is within normal range." },
  { term: "Poor", definition: "Equipment is not performing well and needs further improvement. Life expectancy will be reduced by 25%." },
  { term: "BRAND", definition: "Brand of the equipment." },
  { term: "CAPACITY", definition: "Capacity of the equipment as indicated on the faceplate or technical specifications manual." },
  { term: "EQPT CODE", definition: "Code of the equipment indicating group, location, sequence number, tag, and serviced area at one glance." },
  { term: "EQPT GROUP", definition: "As indicated in the MEPF Life Expectancy Criteria." },
  { term: "EQPT LOCATION", definition: "Location of the MEPF equipment." },
  { term: "EQPT SEQUENCE NO.", definition: "Sequence number of the equipment when more than one equipment item in the same location serves the same area." },
  { term: "EQPT TAG", definition: "Equipment type as abbreviated by the property." },
  { term: "EQPT TYPE", definition: "Type of MEPF equipment." },
  { term: "HOTEL ASSET CODE", definition: "As verified with Finance (Fixed Asset)." },
  { term: "INSTALLATION DATE", definition: "Date of first installation." },
  { term: "NORMAL LIFE EXPECTANCY", definition: "As indicated in the MEPF Life Expectancy Criteria." },
  { term: "REMARKS", definition: "Remarks for any ongoing or past major repairs or rehabilitation done by the property." },
  { term: "REPLACEMENT DATE", definition: "Scheduled replacement based on the installation date, normal life, and adjusted life expectancy." },
  { term: "SERVICED AREA", definition: "Area that the equipment is servicing or the person responsible for the tools and instruments." },
] satisfies TermDefinition[];

export type IndexEntry = {
  id: string;
  number: string;
  equipmentGroup: string;
  description: string;
};

export const indexEntries: IndexEntry[] = [
  { id: "index-1", number: "1", equipmentGroup: "HVAC-1", description: "HEATING, VENTILATION AND AIR-CONDITIONING" },
  { id: "index-1-detail-1", number: "", equipmentGroup: "", description: "Chiller & Cooling Tower" },
  { id: "index-2", number: "2", equipmentGroup: "HVAC-2", description: "HEATING, VENTILATION AND AIR-CONDITIONING" },
  { id: "index-2-detail-1", number: "", equipmentGroup: "", description: "Air Supply Equipment" },
  { id: "index-2-detail-2", number: "", equipmentGroup: "", description: "VFD of Air Supply Equipment" },
  { id: "index-3", number: "3", equipmentGroup: "HVAC-3", description: "HEATING, VENTILATION AND AIR-CONDITIONING" },
  { id: "index-3-detail-1", number: "", equipmentGroup: "", description: "Air Exhaust Equipment" },
  { id: "index-3-detail-2", number: "", equipmentGroup: "", description: "VFD of Air Exhaust Equipment" },
  { id: "index-4", number: "4", equipmentGroup: "HVAC-4", description: "HEATING, VENTILATION AND AIR-CONDITIONING" },
  { id: "index-4-detail-1", number: "", equipmentGroup: "", description: "Fan Coil Unit" },
  { id: "index-5", number: "5", equipmentGroup: "IS", description: "ICE-STORAGE SYSTEM" },
  { id: "index-6", number: "6", equipmentGroup: "BLR", description: "BOILER AND AUXILIARY" },
  { id: "index-6-detail-1", number: "", equipmentGroup: "", description: "Boiler & Auxiliary Equipment" },
  { id: "index-6-detail-2", number: "", equipmentGroup: "", description: "Electrical Control Equipment of Boiler & Auxiliary Equipment" },
  { id: "index-7", number: "7", equipmentGroup: "PS-1", description: "PUMP SYSTEM - 1" },
  { id: "index-7-detail-1", number: "", equipmentGroup: "", description: "Pump of Heating & Air-conditioning System" },
  { id: "index-7-detail-2", number: "", equipmentGroup: "", description: "Electrical Control Equipment of Pump of Heating & Air-conditioning System" },
  { id: "index-7-detail-3", number: "", equipmentGroup: "", description: "VFD of Pump of Heating & Air-conditioning System" },
  { id: "index-8", number: "8", equipmentGroup: "PS-2", description: "PUMP SYSTEM - 2" },
  { id: "index-8-detail-1", number: "", equipmentGroup: "", description: "Pump of Hot Water System" },
  { id: "index-8-detail-2", number: "", equipmentGroup: "", description: "Electrical Control Equipment of Pump of Hot Water System" },
  { id: "index-8-detail-3", number: "", equipmentGroup: "", description: "VFD of Pump of Hot Water System" },
  { id: "index-9", number: "9", equipmentGroup: "PS-3", description: "PUMP SYSTEM - 3" },
  { id: "index-9-detail-1", number: "", equipmentGroup: "", description: "Pump of Potable/Flushing/Make-up Water System" },
  { id: "index-9-detail-2", number: "", equipmentGroup: "", description: "Electrical Control Equipment of Potable/Flushing/Make-up Water System" },
];