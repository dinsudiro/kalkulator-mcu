'use strict';

const config = window.MCU_AUTH;
const loginForm = document.getElementById('loginForm');
const loginStatus = document.getElementById('loginStatus');
let session;
let expiryTimer;

function configured() {
 try {
  const url = new URL(config.projectUrl);
  return url.protocol === 'https:' && url.origin === config.projectUrl &&
   /^sb_publishable_[A-Za-z0-9_-]+$/.test(config.publishableKey);
 } catch { return false; }
}

async function request(path, {token, body, method = 'GET'} = {}) {
 const headers = {apikey:config.publishableKey};
 if(token) headers.Authorization = 'Bearer '+token;
 if(body) headers['Content-Type'] = 'application/json';
 return fetch(config.projectUrl+path, {method,headers,body:body && JSON.stringify(body),cache:'no-store'});
}

function lock(message) {
 clearTimeout(expiryTimer);
 session = undefined;
 data = undefined;
 defaults = undefined;
 snapshot = undefined;
 selected.clear();
 for(const id of ['units','ops','picker','savedList','breakdown']) document.getElementById(id).replaceChildren();
 document.getElementById('protectedApp').hidden = true;
 document.getElementById('loginScreen').hidden = false;
 loginStatus.textContent = message;
}

loginForm.onsubmit = async event => {
 event.preventDefault();
 const button = loginForm.querySelector('button');
 button.disabled = true;
 loginStatus.textContent = 'Memeriksa akun…';
 let credentials;
 try {
  const response = await request('/auth/v1/token?grant_type=password',{
   method:'POST',body:{email:loginForm.elements.email.value.trim(),password:loginForm.elements.password.value}
  });
  if(!response.ok) throw new Error(response.status === 429 ? 'Terlalu banyak percobaan. Coba lagi nanti.' : 'Login gagal. Periksa email dan kata sandi.');
  credentials = await response.json();
  const references = await request('/rest/v1/mcu_reference?id=eq.1&select=payload',{token:credentials.access_token});
  if(!references.ok) throw new Error('Referensi biaya belum bisa diakses.');
  const [row] = await references.json();
  if(!row) throw new Error('Akun belum memiliki akses kalkulator. Hubungi admin.');
  if(!Array.isArray(row.payload?.units) || !Array.isArray(row.payload?.ops)) throw new Error('Referensi biaya belum dikonfigurasi.');
  initializeCalculator(row.payload,credentials.user.id);
  session = credentials.access_token;
  expiryTimer = setTimeout(() => lock('Sesi berakhir. Silakan masuk kembali.'),Number(credentials.expires_in)*1000);
  loginForm.reset();
  document.getElementById('loginScreen').hidden = true;
  document.getElementById('protectedApp').hidden = false;
 } catch(error) {
  loginStatus.textContent = error instanceof TypeError ? 'Layanan login tidak dapat dihubungi. Coba lagi.' : error.message;
 } finally {
  loginForm.elements.password.value = '';
  button.disabled = false;
 }
};

document.getElementById('logout').onclick = () => {
 const token = session;
 lock('Anda sudah keluar.');
 if(token) request('/auth/v1/logout?scope=local',{method:'POST',token}).catch(()=>{});
};

if(!configured()) {
 loginForm.querySelector('button').disabled = true;
 loginStatus.textContent = 'Layanan login belum dikonfigurasi. Hubungi admin.';
}
