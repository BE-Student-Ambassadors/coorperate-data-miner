import { initializeApp } from 'firebase/app'
import { getAuth } from 'firebase/auth'
import { getFirestore } from 'firebase/firestore'

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY || '',
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN || 'dr-system-c7488.firebaseapp.com',
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID || 'dr-system-c7488',
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET || 'dr-system-c7488.firebasestorage.app',
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID || '886413264215',
  appId: import.meta.env.VITE_FIREBASE_APP_ID || '1:886413264215:web:044c9d88a98d7373c6aa12',
  measurementId: import.meta.env.VITE_FIREBASE_MEASUREMENT_ID || 'G-79M3537BGX',
}

export const app = initializeApp(firebaseConfig)
export const auth = getAuth(app)
export const db = getFirestore(app)

