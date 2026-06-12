// ============================================================
//  FIREBASE CONFIGURATION
//  Yahan apna Firebase config paste karein
//  (See SETUP_GUIDE.md for step-by-step instructions)
// ============================================================

import { initializeApp } from 'firebase/app';
import { getDatabase, ref, onValue, set, get } from 'firebase/database';

// 🔥 STEP 1: Replace these values with YOUR Firebase project config
// Get this from: Firebase Console → Project Settings → Your Apps → SDK setup
const firebaseConfig = {
  apiKey:            "PASTE_YOUR_API_KEY_HERE",
  authDomain:        "PASTE_YOUR_AUTH_DOMAIN_HERE",
  databaseURL:       "PASTE_YOUR_DATABASE_URL_HERE",
  projectId:         "PASTE_YOUR_PROJECT_ID_HERE",
  storageBucket:     "PASTE_YOUR_STORAGE_BUCKET_HERE",
  messagingSenderId: "PASTE_YOUR_MESSAGING_SENDER_ID_HERE",
  appId:             "PASTE_YOUR_APP_ID_HERE",
};

// Initialize Firebase
const app  = initializeApp(firebaseConfig);
const db   = getDatabase(app);

// ── Database helpers ────────────────────────────────────────

/** Load all app data once */
export async function loadData() {
  try {
    const snapshot = await get(ref(db, 'kns'));
    if (snapshot.exists()) return snapshot.val();
    return null;
  } catch (e) {
    console.error('Firebase load error:', e);
    return null;
  }
}

/** Save all app data (overwrites) */
export async function saveData(data) {
  try {
    await set(ref(db, 'kns'), data);
  } catch (e) {
    console.error('Firebase save error:', e);
  }
}

/** Listen for real-time changes from other devices */
export function listenData(callback) {
  return onValue(ref(db, 'kns'), (snapshot) => {
    if (snapshot.exists()) callback(snapshot.val());
  });
}

export { db };
