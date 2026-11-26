import { initializeApp, getApps, getApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: 'AIzaSyD_Ioqa_faTECheAIS6c-5VCFqqRrLKNE',
  authDomain: 'inventario-loja-fe612.firebaseapp.com',
  projectId: 'inventario-loja-fe612',
  storageBucket: 'inventario-loja-fe612.firebasestorage.app',
  messagingSenderId: '855356024098',
  appId: '1:855356024098:web:7ee56e07eb002a4f43717b',
};

const app = !getApps().length ? initializeApp(firebaseConfig) : getApp();
export const db = getFirestore(app);
