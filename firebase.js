import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
import { getAuth, signInAnonymously } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
import { getFirestore, collection, onSnapshot, addDoc, updateDoc, deleteDoc, doc } 
  from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

// make helpers visible globally
window.onSnapshot = onSnapshot;
window.collection = collection;
window.doc = doc;
window.addDoc = addDoc;
window.updateDoc = updateDoc;
window.deleteDoc = deleteDoc;

const firebaseConfig = {
  apiKey: "AIzaSyDEGYeA0ere_txZPbwxMH5-BRflZqh_ef0",
  authDomain: "wikitehra.firebaseapp.com",
  projectId: "wikitehra",
  storageBucket: "wikitehra.appspot.com",
  messagingSenderId: "761691537990",
  appId: "1:761691537990:web:3da838e3b77bf2d052c047"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const auth = getAuth(app);

// Reference to 'branches' collection
const branchesRef = collection(db, "monitoring", "_meta", "branches");

// Reference to 'progress' collection (months)
const progressRef = (month, year) => {
    return collection(db, "monitoring", "_meta", "progress", `${year}-${month}`, "entries");
};

// expose db and auth globally
window.db = db;
window.auth = auth;

let monitoringCollectionRef;

async function initFirebase() {
  try {
    await signInAnonymously(auth);
    const appId = typeof __app_id !== 'undefined' ? __app_id : 'sched-master-default';
    
    // Use the correct path for 'branches'
    monitoringCollectionRef = collection(db, "monitoring", "_meta", "branches");
    window.monitoringCollectionRef = monitoringCollectionRef;
    
    console.log("🔥 Firebase connected for:", appId);
    listenForMonitoringUpdates();
  } catch (error) {
    console.error("Firebase initialization failed:", error);
    if (typeof showWarning === "function")
      showWarning("Could not connect to real-time monitoring service.");
  }
}

window.addEventListener("load", () => setTimeout(initFirebase, 500));
