// firebase.ts
import { initializeApp } from "firebase/app";
import {
  getFirestore,
  collection,
  addDoc,
  getDocs,
  updateDoc,
  doc,
  onSnapshot,
  deleteDoc,
  setDoc
} from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDEGYeA0ere_txZPbwxMH5-BRflZqh_ef0",
  authDomain: "wikitehra.firebaseapp.com",
  projectId: "wikitehra",
  storageBucket: "wikitehra.firebasestorage.app",
  messagingSenderId: "761691537990",
  appId: "1:761691537990:web:70c47b4627350ade52c047"
};

// Initialize Firebase once
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);

// Reference to a collection helper (keeps original functions available)
const myCollection = (name: string) => collection(db, name);

// --- FUNCTIONS (exported) ---

export async function addData(collectionName: string, data: any) {
  try {
    const colRef = myCollection(collectionName);
    const docRef = await addDoc(colRef, data);
    console.log("Document added with ID:", docRef.id);
    return docRef.id;
  } catch (error) {
    console.error("Error adding document:", error);
  }
}

export async function updateData(collectionName: string, docId: string, newData: any) {
  try {
    const docRef = doc(db, collectionName, docId);
    await updateDoc(docRef, newData);
    console.log("Document updated:", docId);
  } catch (error) {
    console.error("Error updating document:", error);
  }
}

export async function setData(collectionName: string, docId: string, newData: any, merge = false) {
  try {
    const docRef = doc(db, collectionName, docId);
    if (merge) await setDoc(docRef, newData, { merge: true });
    else await setDoc(docRef, newData);
    console.log("Document set (overwritten):", docId);
  } catch (error) {
    console.error("Error setting document:", error);
  }
}

export async function deleteData(collectionName: string, docId: string) {
  try {
    const docRef = doc(db, collectionName, docId);
    await deleteDoc(docRef);
    console.log("Document deleted:", docId);
  } catch (error) {
    console.error("Error deleting document:", error);
  }
}

export async function getAllData(collectionName: string) {
  try {
    const colRef = myCollection(collectionName);
    const snapshot = await getDocs(colRef);
    const results: any[] = [];
    snapshot.forEach(d => results.push({ id: d.id, ...d.data() }));
    return results;
  } catch (error) {
    console.error("Error getting documents:", error);
    return [];
  }
}

// Real-time listener for a collection
export function subscribeToData(collectionName: string, callback: (data: any[]) => void) {
  const colRef = myCollection(collectionName);
  return onSnapshot(colRef, (snapshot) => {
    const results: any[] = [];
    snapshot.forEach(doc => results.push({ id: doc.id, ...doc.data() }));
    callback(results);
  }, (error) => {
    console.error("Realtime snapshot error:", error);
  });
}

// Real-time listener for a single doc (useful for sharedMonitoringData)
export function subscribeToDoc(collectionName: string, docId: string, onNext: (snap: any) => void, onError?: (e: any) => void) {
  const docRef = doc(db, collectionName, docId);
  return onSnapshot(docRef, onNext, onError);
}

// Convenience helper to set a shared doc
export async function setSharedDoc(collectionName: string, docId: string, payload: any, merge = true) {
  try {
    await setData(collectionName, docId, payload, merge);
  } catch (e) {
    console.warn("setSharedDoc failed:", e);
  }
}
