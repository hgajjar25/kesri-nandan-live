import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

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
export default app;
