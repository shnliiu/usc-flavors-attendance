import { initializeApp } from "https://www.gstatic.com/firebasejs/12.6.0/firebase-app.js";
import {
  getFirestore, doc, getDoc, setDoc, updateDoc, serverTimestamp,
  onSnapshot, collection, getDocs, writeBatch
} from "https://www.gstatic.com/firebasejs/12.6.0/firebase-firestore.js";
import {
  getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/12.6.0/firebase-auth.js";
import { firebaseConfig } from "./firebase-config.js";

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);
const provider = new GoogleAuthProvider();

export async function signInOfficer() {
  const result = await signInWithPopup(auth, provider);
  const officer = await getDoc(doc(db, "officers", result.user.uid));
  if (!officer.exists()) {
    await signOut(auth);
    throw new Error("This Google account is not authorized as a Flavors officer.");
  }
  return result.user;
}

export function watchAuth(callback) {
  return onAuthStateChanged(auth, callback);
}

export async function signOutOfficer() {
  await signOut(auth);
}

export async function getPublicMember(code) {
  const snap = await getDoc(doc(db, "publicMembers", code));
  return snap.exists() ? { found: true, ...snap.data() } : { found: false };
}

export async function getCurrentMeeting() {
  const snap = await getDoc(doc(db, "publicState", "current"));
  return snap.exists() ? snap.data() : { open: false };
}

export function watchAttendance(meetingId, code, callback) {
  return onSnapshot(doc(db, "meetings", meetingId, "attendance", code), snap => {
    callback(snap.exists() ? snap.data() : null);
  });
}

export async function checkInMember(meetingId, code) {
  const privateMember = await getDoc(doc(db, "members", code));
  if (!privateMember.exists() || privateMember.data().active === false) {
    return { type: "notfound" };
  }
  const ref = doc(db, "meetings", meetingId, "attendance", code);
  const existing = await getDoc(ref);
  if (existing.exists() && existing.data().status === "Present") {
    return { type: "already", name: privateMember.data().name };
  }
  await setDoc(ref, {
    status: "Present",
    checkedInAt: serverTimestamp()
  }, { merge: true });
  return { type: "success", name: privateMember.data().name };
}

export async function startMeeting(meetingId, label) {
  await setDoc(doc(db, "meetings", meetingId), {
    open: true,
    label,
    openedAt: serverTimestamp()
  }, { merge: true });
  await setDoc(doc(db, "publicState", "current"), {
    open: true,
    meetingId,
    label
  });
}

export async function closeMeeting(meetingId) {
  await updateDoc(doc(db, "meetings", meetingId), {
    open: false,
    closedAt: serverTimestamp()
  });
  await setDoc(doc(db, "publicState", "current"), { open: false });
}

export async function getRosterCount() {
  const snap = await getDocs(collection(db, "members"));
  return snap.size;
}
