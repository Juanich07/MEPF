import { getApp, getApps, initializeApp } from "firebase/app";
import { getAuth, type User } from "firebase/auth";
import {
  addDoc,
  collection,
  deleteDoc,
  doc,
  getDoc,
  getFirestore,
  getDocs,
  onSnapshot,
  orderBy,
  query,
  serverTimestamp,
  setDoc,
  Timestamp,
  where,
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
export const auth = app ? getAuth(app) : null;

export type AdminUserProfile = {
  uid: string;
  email: string;
  isAdmin: boolean;
  lastSeen: Date;
};

export type AdminActivity = {
  id: string;
  actorUid: string;
  actorEmail: string;
  action: "add" | "edit" | "delete" | "import";
  entity: string;
  entityId: string;
  description: string;
  changedFields: string[];
  createdAt: Date;
};

export type AdminActivityInput = Omit<AdminActivity, "id" | "actorUid" | "actorEmail" | "createdAt">;

const adminUsersCollection = () => {
  if (!database) throw new Error("Firebase is not configured.");
  return collection(database, "adminUsers");
};
const adminActivityCollection = () => {
  if (!database) throw new Error("Firebase is not configured.");
  return collection(database, "adminActivity");
};
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
  const systemMarker = await getDoc(doc(database!, "system", systemMarkerId));
  if (systemMarker.exists()) return;

  const legacySnapshot = await getDocs(legacyCollection);
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
    }
    await batch.commit();
  }

  if (legacyAssets.length === 0 && !legacyMarker) {
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

export async function syncAdminUserProfile(user: User, isAdmin: boolean): Promise<void> {
  if (!database) throw new Error("Firebase is not configured.");
  await setDoc(doc(adminUsersCollection(), user.uid), {
    uid: user.uid,
    email: user.email ?? "",
    isAdmin,
    lastSeen: serverTimestamp(),
  });
}

export async function recordAdminActivity(input: AdminActivityInput): Promise<void> {
  if (!database || !auth?.currentUser) throw new Error("A signed-in Firebase user is required.");
  await addDoc(adminActivityCollection(), {
    ...input,
    description: input.description.slice(0, 200),
    changedFields: input.changedFields.slice(0, 20),
    actorUid: auth.currentUser.uid,
    actorEmail: auth.currentUser.email ?? "",
    createdAt: serverTimestamp(),
  });
}

export function watchAdminDashboard(
  onUsers: (users: AdminUserProfile[]) => void,
  onError: (error: Error) => void,
) {
  if (!database) throw new Error("Firebase is not configured.");
  return onSnapshot(
    query(adminUsersCollection(), orderBy("lastSeen", "desc")),
    (snapshot) => {
      try {
        const users = snapshot.docs.map((userDocument) => {
          const data = userDocument.data();
          if (typeof data.email !== "string" || typeof data.isAdmin !== "boolean" || !(data.lastSeen instanceof Timestamp)) {
            throw new Error(`User profile "${userDocument.id}" is invalid.`);
          }
          return {
            uid: userDocument.id,
            email: data.email,
            isAdmin: data.isAdmin,
            lastSeen: data.lastSeen.toDate(),
          };
        });
        onUsers(users);
      } catch (error) {
        onError(error instanceof Error ? error : new Error(String(error)));
      }
    },
    onError,
  );
}

export function watchAdminActivitiesForUser(
  userId: string,
  onActivities: (activities: AdminActivity[]) => void,
  onError: (error: Error) => void,
) {
  if (!database) throw new Error("Firebase is not configured.");
  return onSnapshot(
    query(adminActivityCollection(), where("actorUid", "==", userId)),
    (snapshot) => {
      try {
        const activities = snapshot.docs.map((activityDocument) => {
          const data = activityDocument.data();
          if (
            typeof data.actorUid !== "string"
            || typeof data.actorEmail !== "string"
            || !["add", "edit", "delete", "import"].includes(data.action)
            || typeof data.entity !== "string"
            || typeof data.entityId !== "string"
            || typeof data.description !== "string"
            || !Array.isArray(data.changedFields)
            || !(data.createdAt instanceof Timestamp)
          ) {
            throw new Error(`Activity record "${activityDocument.id}" is invalid.`);
          }
          return {
            id: activityDocument.id,
            actorUid: data.actorUid,
            actorEmail: data.actorEmail,
            action: data.action as AdminActivity["action"],
            entity: data.entity,
            entityId: data.entityId,
            description: data.description,
            changedFields: data.changedFields.filter((field): field is string => typeof field === "string"),
            createdAt: data.createdAt.toDate(),
          };
        });
        activities.sort((first, second) => second.createdAt.getTime() - first.createdAt.getTime());
        onActivities(activities);
      } catch (error) {
        onError(error instanceof Error ? error : new Error(String(error)));
      }
    },
    onError,
  );
}

function referenceEntriesCollection(id: string) {
  if (!database) throw new Error("Firebase is not configured.");
  return collection(database, "system", id, "entries");
}

function referenceEntryId<T extends Record<string, unknown>>(item: T): string {
  if (typeof item === "object" && item !== null && "id" in item && typeof item.id === "string") {
    return item.id;
  }
  if (typeof item === "object" && item !== null && "term" in item && typeof item.term === "string") {
    return encodeURIComponent(item.term);
  }
  throw new Error("Reference entry is missing its document ID.");
}

function watchReferenceData<T extends Record<string, unknown>>(
  id: string,
  defaults: T[],
  onData: (items: T[]) => void,
  onError: (error: Error) => void,
  fromFirestore: (data: Record<string, unknown>, id: string) => T,
) {
  if (!database) throw new Error("Firebase is not configured.");
  let active = true;
  const entries = referenceEntriesCollection(id);

  void (async () => {
    const [existingEntries, legacyDocument] = await Promise.all([
      getDocs(entries),
      getDoc(doc(database!, "system", id)),
    ]);
    if (!active || legacyDocument.data()?.migrationComplete === true) return;
    if (!existingEntries.empty) {
      await setDoc(doc(database!, "system", id), { migrationComplete: true }, { merge: true });
      return;
    }

    const legacyItems = legacyDocument.exists() ? legacyDocument.data().items : undefined;
    if (legacyItems !== undefined && !Array.isArray(legacyItems)) {
      onError(new Error(`Reference data document "${id}" is invalid.`));
      return;
    }

    const initialItems = Array.isArray(legacyItems) ? legacyItems as T[] : defaults;
    if (initialItems.length > 0) {
      const batch = writeBatch(database!);
      for (const item of initialItems) {
        batch.set(doc(entries, referenceEntryId(item)), { ...item });
      }
      await batch.commit();
    }
    await setDoc(doc(database!, "system", id), { migrationComplete: true }, { merge: true });
  })().catch((error: unknown) => {
    if (active) onError(error instanceof Error ? error : new Error(String(error)));
  });

  const unsubscribe = onSnapshot(entries, (snapshot) => {
    try {
      onData(snapshot.docs.map((entry) => fromFirestore(entry.data(), entry.id)));
    } catch (error) {
      onError(error instanceof Error ? error : new Error(String(error)));
    }
  }, onError);
  return () => {
    active = false;
    unsubscribe();
  };
}

export function watchIndexEntriesInFirestore(
  onEntries: (entries: IndexEntry[]) => void,
  onError: (error: Error) => void,
) {
  return watchReferenceData(indexDataDocumentId, indexEntries, onEntries, onError, (data, id) => {
    if (typeof data.number !== "string" || typeof data.equipmentGroup !== "string" || typeof data.description !== "string") {
      throw new Error(`Index entry "${id}" is invalid.`);
    }
    return { id, number: data.number, equipmentGroup: data.equipmentGroup, description: data.description };
  });
}

export function watchTermDefinitionsInFirestore(
  onDefinitions: (definitions: TermDefinition[]) => void,
  onError: (error: Error) => void,
) {
  return watchReferenceData(definitionsDataDocumentId, termDefinitions, onDefinitions, onError, (data, id) => {
    if (typeof data.term !== "string" || typeof data.definition !== "string") {
      throw new Error(`Definition "${id}" is invalid.`);
    }
    return { id, term: data.term, definition: data.definition };
  });
}

export async function saveIndexEntriesToFirestore(entries: IndexEntry[]): Promise<void> {
  await saveReferenceEntries(indexDataDocumentId, entries);
}

export async function saveTermDefinitionsToFirestore(definitions: TermDefinition[]): Promise<void> {
  await saveReferenceEntries(definitionsDataDocumentId, definitions);
}

async function saveReferenceEntries<T extends Record<string, unknown>>(id: string, items: T[]): Promise<void> {
  if (!database) throw new Error("Firebase is not configured.");
  const entries = referenceEntriesCollection(id);
  const snapshot = await getDocs(entries);
  const ids = new Set(items.map(referenceEntryId));
  const batch = writeBatch(database);
  for (const item of items) {
    batch.set(doc(entries, referenceEntryId(item)), { ...item });
  }
  for (const existing of snapshot.docs) {
    if (!ids.has(existing.id)) batch.delete(existing.ref);
  }
  await batch.commit();
}