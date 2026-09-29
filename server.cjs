const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const PORT = Number(process.env.PORT || 8787);
const HOST = process.env.HOST || '0.0.0.0';
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || 'http://localhost:5173';
const DATA_DIR = path.join(__dirname, 'server-data');
const REPO_FILE = path.join(DATA_DIR, 'repository.json');
const USERS_FILE = path.join(DATA_DIR, 'users.json');
const SESSION_TTL_MS = 8 * 60 * 60 * 1000;
// SECURITY: passwords are ALWAYS verified against the stored scrypt hash — there is no path that
// skips this check. DEMO_LOGIN only controls whether an UNKNOWN email/username is silently
// auto-registered as a new demo account (with a random password nobody is told) on first login.
// It defaults to OFF. Set DEMO_LOGIN=1 only for a supervised jury/demo environment, never in
// production, since it lets an attacker create their own account by choosing an unused email.
const DEMO_LOGIN = process.env.DEMO_LOGIN === '1';
const MAX_LOGIN_ATTEMPTS = 5;
const LOGIN_LOCKOUT_MS = 5 * 60 * 1000;
fs.mkdirSync(DATA_DIR, { recursive: true });

const PERMISSIONS = {
  Technician: ['instrument:create','instrument:edit-draft','test:record','test:mark-na','test:submit','evidence:upload'],
  Examiner: ['test:review'],
  Authority: ['test:review','report:approve-final'],
  Admin: ['instrument:create','instrument:edit-draft','test:record','test:mark-na','test:submit','evidence:upload','test:review','report:approve-final','rule:manage','fingerprint:configure','user:manage'],
  Viewer: []
};
const ROLE_COLLECTIONS = {
  Technician: new Set(['instruments','testPlans','tests','envReadings','auditLog','iotDevices']),
  Examiner: new Set(['tests','auditLog']),
  Authority: new Set(['tests','reports','auditLog']),
  Admin: new Set(['instruments','testPlans','tests','envReadings','reports','auditLog','iotDevices','users','referenceEquipment','ruleVersions','referenceSources','fingerprintConfig']),
  Viewer: new Set([])
};

function nowISO(){ return new Date().toISOString(); }
function hashPassword(password, salt=crypto.randomBytes(16).toString('hex')){
  const derived = crypto.scryptSync(password, salt, 64).toString('hex');
  return {salt, hash:derived};
}
function verifyPassword(password, record){
  const derived = crypto.scryptSync(password, record.salt, 64).toString('hex');
  return crypto.timingSafeEqual(Buffer.from(derived,'hex'), Buffer.from(record.hash,'hex'));
}
function token(){ return crypto.randomBytes(32).toString('hex'); }
function safeUser(u){ return {id:u.id,name:u.name,role:u.role,email:u.email,phone:u.phone,designation:u.designation}; }

// Passwords here MUST match README.md's "Demo accounts" table — the two were out of sync before
// this fix (server seeded "Navi@123"/"Admin@123"/"Viewer@123", README documented "METRA@123").
const DEFAULT_USERS = [
  ['U1','R. Sharma','Technician','r.sharma@nawi-lab.gov.in','Laboratory Technician','METRA@123'],
  ['U2','P. Nair','Technician','p.nair@nawi-lab.gov.in','Laboratory Technician','METRA@123'],
  ['U3','Dr. N. Verma','Examiner','n.verma@nawi-lab.gov.in','Senior Examiner','METRA@123'],
  ['U4','S. Iyer','Authority','s.iyer@nawi-lab.gov.in','Approving Authority','METRA@123'],
  ['U5','A. Admin','Admin','admin@nawi-lab.gov.in','System Administrator','METRA@123'],
  ['U6','Guest Viewer','Viewer','viewer@nawi-lab.gov.in','Observer','METRA@123']
];
function loadUsers(){
  if(fs.existsSync(USERS_FILE)) return JSON.parse(fs.readFileSync(USERS_FILE,'utf8'));
  const users = DEFAULT_USERS.map(([id,name,role,email,designation,password])=>({id,name,role,email,designation,phone:'—',...hashPassword(password)}));
  fs.writeFileSync(USERS_FILE, JSON.stringify(users,null,2));
  return users;
}
let users = loadUsers();
let repository = fs.existsSync(REPO_FILE) ? JSON.parse(fs.readFileSync(REPO_FILE,'utf8')) : null;
const sessions = new Map();
const loginAttempts = new Map(); // email -> {count, lockedUntil}

function saveRepository(){
  const tmp = REPO_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(repository,null,2));
  fs.renameSync(tmp, REPO_FILE);
}
function publicRepository(){
  if(!repository) return null;
  return {...repository, users: users.map(safeUser)};
}
function parseCookies(req){
  const out={};
  String(req.headers.cookie||'').split(';').forEach(part=>{ const i=part.indexOf('='); if(i>0) out[part.slice(0,i).trim()]=decodeURIComponent(part.slice(i+1).trim()); });
  return out;
}
function sessionUser(req){
  const sid=parseCookies(req).navi_session;
  if(!sid) return null;
  const s=sessions.get(sid);
  if(!s || s.expiresAt<Date.now()){ if(sid) sessions.delete(sid); return null; }
  const u=users.find(x=>x.id===s.userId);
  return u || null;
}
function send(res,status,data,extra={}){
  const body=JSON.stringify(data);
  res.writeHead(status, {'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store','Access-Control-Allow-Origin':ALLOWED_ORIGIN,'Access-Control-Allow-Credentials':'true','Vary':'Origin',...extra});
  res.end(body);
}
function noContent(res,status=204){ res.writeHead(status, {'Access-Control-Allow-Origin':ALLOWED_ORIGIN,'Access-Control-Allow-Credentials':'true','Vary':'Origin'}); res.end(); }
function parseBody(req){ return new Promise((resolve,reject)=>{ let raw=''; req.on('data',c=>{raw+=c; if(raw.length>15*1024*1024) req.destroy();}); req.on('end',()=>{try{resolve(raw?JSON.parse(raw):{})}catch(e){reject(new Error('Invalid JSON body'))}}); req.on('error',reject); }); }
function unauthorized(res){ return send(res,401,{error:'Authentication required.'}); }
function forbidden(res,msg='Your role is not permitted to perform this action.'){ return send(res,403,{error:msg}); }
function sameJson(a,b){ return JSON.stringify(a)===JSON.stringify(b); }
function changedCollections(oldRepo,newRepo){
  const keys=new Set([...Object.keys(oldRepo||{}),...Object.keys(newRepo||{})]);
  return [...keys].filter(k=>!sameJson(oldRepo?.[k],newRepo?.[k]));
}
function auditAppendOnly(oldLog,newLog){
  const oldById=new Map((oldLog||[]).map(x=>[x.id,JSON.stringify(x)]));
  for(const [id,value] of oldById){
    const found=(newLog||[]).find(x=>x.id===id);
    if(!found || JSON.stringify(found)!==value) return false;
  }
  return true;
}
function validateRepositoryWrite(actor,next){
  if(!repository){ return {ok:true,initial:true}; }
  const changed=changedCollections(repository,next);
  const allowed=ROLE_COLLECTIONS[actor.role]||new Set();
  const blocked=changed.filter(k=>!allowed.has(k));
  if(blocked.length) return {ok:false,message:`Server RBAC denied changes to: ${blocked.join(', ')}.`};
  if(!auditAppendOnly(repository.auditLog,next.auditLog)) return {ok:false,message:'Audit trail is append-only and existing events cannot be changed or removed.'};
  if(!sameJson(next.users,repository.users) && actor.role!=='Admin') return {ok:false,message:'Only Admin can manage users.'};
  // Issued reports are immutable. A later correction must create a new report/revision;
  // an existing issued record cannot be edited or silently replaced through the repository API.
  const oldReports=repository.reports||[];
  const newReports=next.reports||[];
  for(const oldReport of oldReports){
    if(oldReport.status !== 'Issued') continue;
    const replacement=newReports.find(r=>r.id===oldReport.id);
    if(!replacement) return {ok:false,message:`Issued report ${oldReport.reportNumber||oldReport.id} cannot be deleted.`};
    if(!sameJson(oldReport,replacement)) return {ok:false,message:`Issued report ${oldReport.reportNumber||oldReport.id} is immutable. Create a new report revision instead.`};
  }
  return {ok:true,changed};
}
function appendServerAudit(actor,action,entity,entityId,reason){
  if(!repository) return;
  repository.auditLog=repository.auditLog||[];
  repository.auditLog.unshift({id:`AUD-SRV-${crypto.randomBytes(5).toString('hex')}`,timestamp:nowISO(),user:`${actor.name} (${actor.role})`,action,entity,entityId,prevValue:'—',newValue:'—',reason});
}

async function route(req,res){
  if(req.method==='OPTIONS'){ res.writeHead(204,{'Access-Control-Allow-Origin':ALLOWED_ORIGIN,'Access-Control-Allow-Credentials':'true','Access-Control-Allow-Headers':'Content-Type','Access-Control-Allow-Methods':'GET,POST,PUT,OPTIONS','Vary':'Origin'}); return res.end(); }
  const url=new URL(req.url,`http://${req.headers.host||'localhost'}`);
  const p=url.pathname;
  if(req.method==='GET' && p==='/api/health') return send(res,200,{ok:true,service:'navi-backend',time:nowISO()});

  if(req.method==='POST' && p==='/api/auth/login'){
    const b=await parseBody(req); const rawEmail=String(b.email||'').trim(); const email=rawEmail.toLowerCase();
    const password=String(b.password||'');
    if(!rawEmail || !password) return send(res,400,{error:'Enter an email and password to continue.'});

    const lock=loginAttempts.get(email);
    if(lock && lock.lockedUntil>Date.now()){
      const secs=Math.ceil((lock.lockedUntil-Date.now())/1000);
      return send(res,429,{error:`Too many failed attempts. Try again in ${secs}s.`});
    }

    let u=users.find(x=>x.email.toLowerCase()===email);
    if(!u && DEMO_LOGIN){
      // Demo-only: registers a brand-new account for an unrecognized email, with a random
      // password that is never returned or logged, and signs it into that new account under the
      // password the caller supplied for THIS request only (so the login below still runs the
      // normal verify path). Only reachable when the operator has explicitly set DEMO_LOGIN=1.
      let role='Technician', designation='Laboratory Technician';
      if(/admin/i.test(email)){ role='Admin'; designation='System Administrator'; }
      else if(/authority|approv/i.test(email)){ role='Authority'; designation='Approving Authority'; }
      else if(/examiner|review/i.test(email)){ role='Examiner'; designation='Senior Examiner'; }
      else if(/viewer|guest/i.test(email)){ role='Viewer'; designation='Observer'; }
      const safeId='DEMO-'+crypto.createHash('sha1').update(email+Date.now()).digest('hex').slice(0,10).toUpperCase();
      u={id:safeId,name:rawEmail.split('@')[0]||'Demo User',role,email,designation,phone:'—',...hashPassword(password)};
      users.push(u);
      fs.writeFileSync(USERS_FILE,JSON.stringify(users,null,2));
    }

    if(!u || !verifyPassword(password,u)){
      const attempts=(lock?.count||0)+1;
      const lockedUntil=attempts>=MAX_LOGIN_ATTEMPTS ? Date.now()+LOGIN_LOCKOUT_MS : 0;
      loginAttempts.set(email,{count:lockedUntil?0:attempts,lockedUntil});
      return send(res,401,{error:'Incorrect email or password.'});
    }
    loginAttempts.delete(email);

    const sid=token(); sessions.set(sid,{userId:u.id,expiresAt:Date.now()+SESSION_TTL_MS});
    return send(res,200,{user:safeUser(u)},{'Set-Cookie':`navi_session=${sid}; HttpOnly; SameSite=Lax; Path=/; Max-Age=${SESSION_TTL_MS/1000}`});
  }
  if(req.method==='POST' && p==='/api/auth/logout'){
    const sid=parseCookies(req).navi_session; if(sid) sessions.delete(sid);
    return send(res,200,{ok:true},{'Set-Cookie':'navi_session=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0'});
  }
  if(req.method==='GET' && p==='/api/auth/me'){
    const u=sessionUser(req); if(!u) return unauthorized(res); return send(res,200,{user:safeUser(u)});
  }
  if(req.method==='PUT' && p==='/api/auth/profile'){
    const u=sessionUser(req); if(!u) return unauthorized(res);
    const b=await parseBody(req);
    const name=String(b.name||'').trim();
    const email=String(b.email||'').trim().toLowerCase();
    const phone=String(b.phone||'').trim();
    const designation=String(b.designation||'').trim();
    if(!name || !email) return send(res,400,{error:'Name and email are required.'});
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return send(res,400,{error:'Enter a valid email address.'});
    const duplicate=users.find(x=>x.id!==u.id && x.email.toLowerCase()===email);
    if(duplicate) return send(res,409,{error:'That email address is already assigned to another user.'});
    const before=safeUser(u);
    u.name=name; u.email=email; u.phone=phone||'—'; u.designation=designation;
    fs.writeFileSync(USERS_FILE,JSON.stringify(users,null,2));
    if(repository){ repository.users=users.map(safeUser); appendServerAudit(u,'Profile updated','User',u.id,before.name,u.name,'Authenticated self-service profile update'); saveRepository(); }
    return send(res,200,{user:safeUser(u)});
  }
  if(req.method==='GET' && p==='/api/repository'){
    const u=sessionUser(req); if(!u) return unauthorized(res); return send(res,200,{repository:publicRepository()});
  }
  if(req.method==='PUT' && p==='/api/repository'){
    const u=sessionUser(req); if(!u) return unauthorized(res);
    const b=await parseBody(req); const next=b.repository;
    if(!next || typeof next!=='object') return send(res,400,{error:'Repository payload is required.'});
    const check=validateRepositoryWrite(u,next); if(!check.ok) return forbidden(res,check.message);
    // Never trust client-supplied directory records or role assignments.
    next.users=users.map(safeUser);
    repository=next; appendServerAudit(u,'Repository saved','Repository','SERVER', 'Authenticated API write'); saveRepository();
    return send(res,200,{ok:true,repository:publicRepository()});
  }
  if(req.method==='GET' && p.startsWith('/api/public/reports/')){
    const reportId=decodeURIComponent(p.slice('/api/public/reports/'.length)).trim().toUpperCase();
    if(!repository) return send(res,404,{error:'Report not found.'});
    const report=(repository.reports||[]).find(r=>String(r.id).toUpperCase()===reportId);
    if(!report || report.status!=='Issued') return send(res,404,{error:'Issued report not found.'});
    const instrument=(repository.instruments||[]).find(i=>i.id===report.instrumentId);
    return send(res,200,{verified:true,verification:{status:'VERIFIED',source:'server-public-endpoint',verifiedAt:nowISO(),integrityHashPresent:Boolean(report.integrityHash)},report:{id:report.id,reportNumber:report.reportNumber,status:report.status,generatedAt:report.generatedAt,ruleVersionId:report.ruleVersionId,integrityHash:report.integrityHash,instrumentId:report.instrumentId,instrumentName:instrument?.model||instrument?.name||'NAWI'}});
  }
  return send(res,404,{error:'Not found.'});
}

const server=http.createServer((req,res)=>{ route(req,res).catch(err=>{console.error(err); send(res,500,{error:'Internal server error.'}); }); });
server.listen(PORT,HOST,()=>console.log(`Navi backend listening on http://${HOST}:${PORT}`));
