import { getApp, getApps, initializeApp } from "firebase/app";
import {
  collection,
  deleteDoc,
  doc,
  getDoc,
  getFirestore,
  getDocs,
  onSnapshot,
  setDoc,
  writeBatch,
} from "firebase/firestore";
import { demoAssets, type Asset } from "@/lib/assets";
import { defaultSheets, indexEntries, termDefinitions } from "@/lib/sheets";
import type { IndexEntry, TermDefinition, WorkbookSheet } from "@/lib/sheets";

const firebaseConfig = {
  apiKey: process.env.NEXT_PUBLIC_FIREBASE_API_KEY,
  authDomain: process.env.NEXT_PUBLIC_FIREBASE_AUTH_DOMAIN,
  projectId: process.env.NEXT_PUBLIC_FIREBASE_PROJECT_ID,
  storageBucket: process.env.NEXT_PUBLIC_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: process.env.NEXT_PUBLIC_FIREBASE_MESSAGING_SENDER_ID,
  appId: process.env.NEXT_PUBLIC_FIREBASE_APP_ID,
};

const initializationMarkerId = "mepf-initialized-marker";
const systemMarkerId = "assets-initialized";
const indexDataDocumentId = "reference-index";
const definitionsDataDocumentId = "reference-definitions";

export const firebaseConfigured = Object.values(firebaseConfig).every(Boolean);

const app = firebaseConfigured
  ? getApps().length > 0
    ? getApp()
    : initializeApp(firebaseConfig)
  : null;

const database = app ? getFirestore(app) : null;
const legacyAssetsCollection = () => {
  if (!database) throw new Error("Firebase is not configured.");
  return collection(database, "assets");
};
const workbookPagesCollection = () => {
  if (!database) throw new Error("Firebase is not configured.");
  return collection(database, "workbookPages");
};
const assetGroupCollection = (group: string) => {
  if (!database) throw new Error("Firebase is not configured.");
  const collectionId = group.trim().replace(/[^a-zA-Z0-9_-]+/g, "-").replace(/^-+|-+$/g, "") || "Unassigned";
  return collection(database, collectionId);
};

function sortAssets(assets: Asset[]) {
  return assets.sort((first, second) =>
    (first.sortOrder ?? Number.MAX_SAFE_INTEGER) - (second.sortOrder ?? Number.MAX_SAFE_INTEGER)
    || first.id.localeCompare(second.id),
  );
}

export async function getAssetsFromFirestore(): Promise<Asset[]> {
  const pagesSnapshot = await getDocs(workbookPagesCollection());
  const groups = [
    ...defaultSheets.filter((sheet) => sheet.kind === "group").map((sheet) => sheet.group ?? sheet.label),
    ...pagesSnapshot.docs.map((pageDocument) => {
      const page = pageDocument.data() as WorkbookSheet;
      return page.group ?? page.label;
    }),
    "Unassigned",
  ];
  const snapshots = await Promise.all([...new Set(groups)].map((group) => getDocs(assetGroupCollection(group))));
  return sortAssets(snapshots.flatMap((snapshot) => snapshot.docs.map((assetDocument) => {
    const data = assetDocument.data() as Omit<Asset, "id">;
    return { ...data, id: assetDocument.id, equipmentGroup: data.equipmentGroup || "Unassigned", remarks: data.remarks ?? "" };
  })));
}

export async function initializeAssetsInFirestore(): Promise<void> {
  const legacyCollection = legacyAssetsCollection();
  const [legacySnapshot, systemMarker] = await Promise.all([
    getDocs(legacyCollection),
    getDoc(doc(database!, "system", systemMarkerId)),
  ]);
  const legacyMarker = legacySnapshot.docs.find((assetDocument) => assetDocument.id === initializationMarkerId);
  const legacyAssets = legacySnapshot.docs.filter((assetDocument) => assetDocument.id !== initializationMarkerId);

  for (let index = 0; index < legacyAssets.length; index += 200) {
    const batch = writeBatch(database!);
    for (const [offset, assetDocument] of legacyAssets.slice(index, index + 200).entries()) {
      const data = assetDocument.data() as Omit<Asset, "id">;
      const group = data.equipmentGroup || "Unassigned";
      const asset: Asset = {
        ...data,
        id: assetDocument.id,
        equipmentGroup: group,
        remarks: typeof data.remarks === "string" && data.remarks !== "Unawarded" ? data.remarks : "",
        sortOrder: typeof data.sortOrder === "number" ? data.sortOrder : index + offset + 1,
      };
      batch.set(doc(assetGroupCollection(group), asset.id), asset);
      batch.delete(assetDocument.ref);
    }
    await batch.commit();
  }

  if (legacyMarker) await deleteDoc(legacyMarker.ref);

  if (!systemMarker.exists() && legacyAssets.length === 0 && !legacyMarker) {
    const pagesSnapshot = await getDocs(workbookPagesCollection());
    const groups = [
      ...defaultSheets.filter((sheet) => sheet.kind === "group").map((sheet) => sheet.group ?? sheet.label),
      ...pagesSnapshot.docs.map((pageDocument) => {
        const page = pageDocument.data() as WorkbookSheet;
        return page.group ?? page.label;
      }),
    ];
    const groupSnapshots = await Promise.all([...new Set(groups)].map((group) => getDocs(assetGroupCollection(group))));
    if (groupSnapshots.every((snapshot) => snapshot.empty)) {
      const seedBatch = writeBatch(database!);
      for (const asset of demoAssets) seedBatch.set(doc(assetGroupCollection(asset.equipmentGroup), asset.id), asset);
      await seedBatch.commit();
    }
  }

  await setDoc(doc(database!, "system", systemMarkerId), { initialized: true });
}

export function watchAssetsInFirestore(
  onAssets: (assets: Asset[]) => void,
  onError: (error: Error) => void,
  groups: string[],
) {
  const uniqueGroups = [...new Set([...groups, "Unassigned"])];
  const loadedGroups = new Set<string>();
  const assetsByGroup = new Map<string, Asset[]>();
  const unsubscribers = uniqueGroups.map((group) => onSnapshot(
    assetGroupCollection(group),
    (snapshot) => {
      assetsByGroup.set(group, snapshot.docs.map((assetDocument) => {
        const data = assetDocument.data() as Omit<Asset, "id">;
        return { ...data, id: assetDocument.id, equipmentGroup: data.equipmentGroup || group, remarks: data.remarks ?? "" };
      }));
      loadedGroups.add(group);
      if (loadedGroups.size === uniqueGroups.length) {
        onAssets(sortAssets([...assetsByGroup.values()].flat()));
      }
    },
    onError,
  ));
  return () => unsubscribers.forEach((unsubscribe) => unsubscribe());
}

export async function saveAssetToFirestore(asset: Asset, previousGroup = asset.equipmentGroup): Promise<void> {
  const previousCollection = assetGroupCollection(previousGroup || "Unassigned");
  const nextCollection = assetGroupCollection(asset.equipmentGroup || "Unassigned");
  if (previousCollection.path !== nextCollection.path) {
    const batch = writeBatch(database!);
    batch.set(doc(nextCollection, asset.id), asset);
    batch.delete(doc(previousCollection, asset.id));
    await batch.commit();
    return;
  }
  await setDoc(doc(nextCollection, asset.id), asset);
}

export async function saveAssetsToFirestore(assets: Asset[], group: string): Promise<void> {
  const groupCollection = assetGroupCollection(group);
  for (let index = 0; index < assets.length; index += 450) {
    const batch = writeBatch(database!);
    for (const asset of assets.slice(index, index + 450)) {
      batch.set(doc(groupCollection, asset.id), asset);
    }
    await batch.commit();
  }
}

export async function deleteAssetFromFirestore(asset: Asset): Promise<void> {
  await deleteDoc(doc(assetGroupCollection(asset.equipmentGroup || "Unassigned"), asset.id));
}

export function watchWorkbookPagesInFirestore(
  onPages: (pages: WorkbookSheet[]) => void,
  onError: (error: Error) => void,
) {
  return onSnapshot(
    workbookPagesCollection(),
    (snapshot) => {
      onPages(snapshot.docs.map((pageDocument) => ({
        ...pageDocument.data(),
        id: pageDocument.id,
      }) as WorkbookSheet));
    },
    onError,
  );
}

export async function saveWorkbookPageToFirestore(page: WorkbookSheet): Promise<void> {
  await setDoc(doc(workbookPagesCollection(), page.id), page);
}

export async function deleteWorkbookPageFromFirestore(id: string): Promise<void> {
  await deleteDoc(doc(workbookPagesCollection(), id));
}

function watchReferenceData<T>(
  id: string,
  defaults: T[],
  onData: (items: T[]) => void,
  onError: (error: Error) => void,
) {
  if (!database) throw new Error("Firebase is not configured.");
  const referenceDocument = doc(database, "system", id);
  return onSnapshot(referenceDocument, (snapshot) => {
    if (snapshot.exists()) {
      const data = snapshot.data().items;
      if (!Array.isArray(data)) {
        onError(new Error(`Reference data document "${id}" is invalid.`));
        return;
      }
      onData(data as T[]);
      return;
    }
    void setDoc(referenceDocument, { items: defaults }).catch(onError);
  }, onError);
}

export function watchIndexEntriesInFirestore(
  onEntries: (entries: IndexEntry[]) => void,
  onError: (error: Error) => void,
) {
  return watchReferenceData(indexDataDocumentId, indexEntries, onEntries, onError);
}

export function watchTermDefinitionsInFirestore(
  onDefinitions: (definitions: TermDefinition[]) => void,
  onError: (error: Error) => void,
) {
  return watchReferenceData(definitionsDataDocumentId, termDefinitions, onDefinitions, onError);
}

export async function saveIndexEntriesToFirestore(entries: IndexEntry[]): Promise<void> {
  if (!database) throw new Error("Firebase is not configured.");
  await setDoc(doc(database, "system", indexDataDocumentId), { items: entries });
}

export async function saveTermDefinitionsToFirestore(definitions: TermDefinition[]): Promise<void> {
  if (!database) throw new Error("Firebase is not configured.");
  await setDoc(doc(database, "system", definitionsDataDocumentId), { items: definitions });
}