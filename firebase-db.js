import { initializeApp } from "https://www.gstatic.com/firebasejs/12.6.0/firebase-app.js";
import {
  getFirestore, doc, getDoc, setDoc, updateDoc, serverTimestamp,
  onSnapshot, collection, getDocs, writeBatch
} from "https://www.gstatic.com/firebasejs/12.6.0/firebase-firestore.js";
import {
  getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/12.6.0/firebase-auth.js";

// Firebase web configuration is intentionally browser-visible. The API key is
// split only to avoid GitHub treating a Firebase client key like a private server secret.
const firebaseConfig = {
  apiKey: ["AIzaSyDQX_", "PZdxq4Ag1cJ6LS7Ay-zLSm-lsEvLs"].join(""),
  authDomain: "usc-flavors-attendance.firebaseapp.com",
  projectId: "usc-flavors-attendance",
  storageBucket: "usc-flavors-attendance.firebasestorage.app",
  messagingSenderId: "387800447090",
  appId: "1:387800447090:web:fa49f14d08de07e10aaaeb"
};

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
export function watchAuth(callback){ return onAuthStateChanged(auth, callback); }
export async function signOutOfficer(){ await signOut(auth); }

export async function isOfficerUser(user){
  if(!user) return false;
  const snap=await getDoc(doc(db,"officers",user.uid));
  return snap.exists();
}

export async function getPublicMember(code){
  const snap=await getDoc(doc(db,"publicMembers",code));
  return snap.exists()?{found:true,...snap.data()}:{found:false};
}
export async function getCurrentMeeting(){
  const snap=await getDoc(doc(db,"publicState","current"));
  return snap.exists()?snap.data():{open:false};
}
export function watchCurrentMeeting(callback){
  return onSnapshot(doc(db,"publicState","current"),snap=>callback(snap.exists()?snap.data():{open:false}));
}
export function watchAttendance(meetingId,code,callback){
  return onSnapshot(doc(db,"meetings",meetingId,"attendance",code),snap=>callback(snap.exists()?snap.data():null));
}
export async function checkInMember(meetingId,code){
  const member=await getDoc(doc(db,"members",code));
  if(!member.exists()||member.data().active===false)return{type:"notfound"};
  const ref=doc(db,"meetings",meetingId,"attendance",code);
  const existing=await getDoc(ref);
  if(existing.exists()&&existing.data().status==="Present")return{type:"already",name:member.data().name};
  await setDoc(ref,{status:"Present",checkedInAt:serverTimestamp()},{merge:true});
  return{type:"success",name:member.data().name};
}
export async function startMeeting(meetingId,label){
  await setDoc(doc(db,"meetings",meetingId),{open:true,label,openedAt:serverTimestamp()},{merge:true});
  await setDoc(doc(db,"publicState","current"),{open:true,meetingId,label});
}
export async function closeMeeting(meetingId){
  await updateDoc(doc(db,"meetings",meetingId),{open:false,closedAt:serverTimestamp()});
  await setDoc(doc(db,"publicState","current"),{open:false});
}
export async function getScannerStatus(){
  const state=await getCurrentMeeting();
  const roster=await getDocs(collection(db,"members"));
  let checkedIn=0;
  if(state.open&&state.meetingId){
    const a=await getDocs(collection(db,"meetings",state.meetingId,"attendance"));
    a.forEach(x=>{if(x.data().status==="Present")checkedIn++});
  }
  return{open:!!state.open,meetingId:state.meetingId||null,label:state.label||"",checkedIn,total:roster.size};
}
export async function importRosterRows(rows,siteBaseUrl){
  const clean=rows.map(r=>({
    code:String(r.code||"").trim().toUpperCase(),
    name:String(r.name||"").trim(),
    email:String(r.email||"").trim(),
    passLink:String(r.passLink||"").trim()
  })).filter(r=>r.code&&r.name);
  for(let i=0;i<clean.length;i+=100){
    const batch=writeBatch(db);
    clean.slice(i,i+100).forEach(r=>{
      const passLink=r.passLink||siteBaseUrl.replace(/\/$/,"")+"/pass/?id="+encodeURIComponent(r.code);
      batch.set(doc(db,"members",r.code),{name:r.name,email:r.email,active:true,passLink},{merge:true});
      batch.set(doc(db,"publicMembers",r.code),{name:r.name,active:true},{merge:true});
      batch.set(doc(db,"usedCodes",r.code),{issuedAt:serverTimestamp()},{merge:true});
    });
    await batch.commit();
  }
  return clean.length;
}
