import { initializeApp } from "firebase/app";
import { getFirestore, collection, getDocs, setDoc, doc } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyBThcFm8T5oYWDgH1UfLkNm1HQXdMSQen0",
  authDomain: "kesri-nandan.firebaseapp.com",
  projectId: "kesri-nandan",
  storageBucket: "kesri-nandan.firebasestorage.app",
  messagingSenderId: "1002171035433",
  appId: "1:1002171035433:web:974b0ba1786572dd32d138"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);

export async function loadData(collectionName) {
  const snapshot = await getDocs(collection(db, collectionName));
  return snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
}

export async function saveData(collectionName, id, data) {
  await setDoc(doc(db, collectionName, id), data);
}

export default app;
