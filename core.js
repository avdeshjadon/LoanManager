// --- core.js ---
// Firebase configuration aur global variables

const firebaseConfig = {
  apiKey: "AIzaSyD0wHBp_Gb-U7eJFjoa3pTBYdgipxMYyzg",
  authDomain: "globalfinanceconsultant-bf13a.firebaseapp.com",
  projectId: "globalfinanceconsultant-bf13a",
  storageBucket: "globalfinanceconsultant-bf13a.firebasestorage.app",
  messagingSenderId: "611257450437",
  appId: "1:611257450437:web:fda4e59be985ef146e55ac",
  measurementId: "G-KWM1STNNZ0",
};

// Initialize Firebase
firebase.initializeApp(firebaseConfig);

// Global Firebase services
const db = firebase.firestore();
const auth = firebase.auth();
const storage = firebase.storage();

// Global state
window.allCustomers = { active: [], settled: [] };
let recentActivities = [];
let currentUser = null;
let activeSortKey = "name";

// NEW: master auth flag for special shortcut login
window.masterAuth = false;