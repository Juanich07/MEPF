export const assetFields = [
  "equipmentGroup",
  "equipmentType",
  "equipmentTag",
  "equipmentLocation",
  "servicedArea",
  "sequenceNumber",
  "equipmentCode",
  "brand",
  "capacity",
  "hotelAssetCode",
  "installationYear",
  "normalLife",
  "presentCondition",
  "adjustedLife",
  "replacementYear",
  "zoning",
  "phasing",
  "remarks",
] as const;

export type AssetField = (typeof assetFields)[number];
export type RemarksColor = "" | "red" | "green";
export type Asset = Record<AssetField, string> & {
  id: string;
  sortOrder?: number;
  remarksColor?: RemarksColor;
};

function makeAsset(
  id: string,
  servicedArea: string,
  capacity: string,
  zoning: string,
  location = "BMT",
): Asset {
  return {
    id,
    sortOrder: Number(id.slice(-2)),
    equipmentGroup: "HVAC-4",
    equipmentType: "Fan Coil Unit",
    equipmentTag: "FCU",
    equipmentLocation: location,
    servicedArea,
    sequenceNumber: "1",
    equipmentCode: `HVAC-4-FCU-${location}-${servicedArea}-1`,
    brand: "",
    capacity,
    hotelAssetCode: "",
    installationYear: "",
    normalLife: "",
    presentCondition: "",
    adjustedLife: "",
    replacementYear: "",
    zoning,
    phasing: "PHASE 1",
    remarks: "",
    remarksColor: "",
  };
}

export const demoAssets: Asset[] = [
  makeAsset("mepf-fcu-01", "Guest Entrance", "2.5 TR", "BF-05"),
  makeAsset("mepf-fcu-02", "CE Room", "0.84 TR", "BF-03"),
  makeAsset("mepf-fcu-03", "Female Sleeping Quarters", "1.23 TR", "BF-03"),
  makeAsset("mepf-fcu-04", "Eng'r Office", "1.5 TR", "BF-03"),
  makeAsset("mepf-fcu-05", "Purchasing", "1.67 TR", "BF-03"),
  makeAsset("mepf-fcu-06", "Hotel CCTV", "2.20 TR", "BF-03"),
  makeAsset("mepf-fcu-07", "Purchasing Manager", "0.84 TR", "BF-03"),
  makeAsset("mepf-fcu-08", "MTG", "0.84 TR", "BF-03"),
  makeAsset("mepf-fcu-09", "Cost Control", "1.67 TR", "BF-03"),
  makeAsset("mepf-fcu-10", "Receiving", "0.84 TR", "BF-03"),
  makeAsset("mepf-fcu-11", "Male CR & Female CR & Shower", "0.84 TR", "BF-03"),
  makeAsset("mepf-fcu-12", "Female Locker", "0.84 TR", "BF-03"),
];