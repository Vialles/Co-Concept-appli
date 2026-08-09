// Configuration Firebase — à remplir avec les identifiants de votre projet
// (Console Firebase > Paramètres du projet > Vos applications > Config SDK)
//
// Étapes avant de lancer l'app :
// 1. Créer un projet sur https://console.firebase.google.com
// 2. Activer "Authentication" → méthode Email/Mot de passe (ou lien magique)
// 3. Activer "Firestore Database" → mode production
// 4. Copier .env.example en .env.local et renseigner les valeurs ci-dessous
//    (elles se trouvent dans les paramètres du projet Firebase)

import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

export const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);
