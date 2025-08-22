/* app.js - blok 3 (finalny frontend)
   - dopieszczony UX: spinner overlay, blokowanie formularza,
   - pokazanie linku do wygenerowanego pliku (jeśli backend zwróci url),
   - zapamiętywanie email w localStorage,
   - drobne poprawki dostępności i komunikaty.
*/
const $ = id => document.getElementById(id);
const refs = {};
['pion','gtInput','kwInput','email','message','themeToggle','themeIcon','themeText','spinnerContainer','gtTags','kwTags','gtList','kwList'].forEach(i => refs[i] = $(i));

let availableGTs = [], availableKWs = [], messageTimeout;

// ========== Utilities ==========
function clearMessage() {
  clearTimeout(messageTimeout);
  if (!refs.message) return;
  refs.message.style.display = 'none';
  refs.message.innerHTML = '';
}
function setMessageAutoClear() {
  messageTimeout = setTimeout(clearMessage, 180000);
}
function showMessage(text, isError=false) {
  if (!refs.message) return;
  refs.message.style.display = 'block';
  refs.message.style.color = isError ? '#b00020' : '#111';
  refs.message.innerHTML = text;
  setMessageAutoClear();
}
function setLoading(on=true) {
  const btn = document.querySelector('button[type="submit"]');
  if (refs.spinnerContainer) refs.spinnerContainer.style.display = on ? 'block' : 'none';
  if (btn) {
    btn.disabled = on;
    if (on) btn.classList.add('sending'); else btn.classList.remove('sending');
  }
  // disable inputs while loading
  ['pion','gtInput','kwInput','email'].forEach(id => {
    const el = $(id); if (el) el.disabled = on;
  });
}
function saveLastEmail(email) {
  try { if (email) localStorage.lastEmail = email; else localStorage.removeItem('lastEmail'); } catch(e){}
}
function loadLastEmail() {
  try { if (localStorage.lastEmail && refs.email) refs.email.value = localStorage.lastEmail; } catch(e){}
}
function createDownloadButton(url, filename) {
  const wrapperId = 'downloadWrapper';
  let wrap = document.getElementById(wrapperId);
  if (!wrap) {
    wrap = document.createElement('div');
    wrap.id = wrapperId;
    wrap.style.marginTop = '12px';
    if (refs.message) refs.message.insertAdjacentElement('afterend', wrap);
    else document.querySelector('.container').appendChild(wrap);
  } else {
    wrap.innerHTML = '';
  }
  const a = document.createElement('a');
  a.href = url;
  a.target = '_blank';
  a.rel = 'noopener noreferrer';
  a.textContent = filename || 'Pobierz plik';
  a.style.display = 'inline-block';
  a.style.padding = '8px 12px';
  a.style.borderRadius = '8px';
  a.style.background = '#F47B20';
  a.style.color = '#fff';
  a.style.fontWeight = '700';
  a.style.textDecoration = 'none';
  a.style.marginTop = '6px';
  wrap.appendChild(a);
}

// ========== Form helpers ==========
function resetFormExceptPion() {
  if (refs.gtTags) refs.gtTags.innerHTML = '';
  if (refs.kwTags) refs.kwTags.innerHTML = '';
  if (refs.gtList) refs.gtList.innerHTML = '';
  if (refs.kwList) refs.kwList.innerHTML = '';
  if (refs.gtInput) refs.gtInput.value = '';
  if (refs.kwInput) refs.kwInput.value = '';
  // don't clear email (we keep it)
}
function toggleTheme() {
  if (!refs.themeToggle) return;
  const dark = document.body.classList.toggle('dark-mode');
  if (refs.themeIcon) refs.themeIcon.textContent = dark ? '☀️' : '🌙';
  if (refs.themeText) refs.themeText.textContent = dark ? 'Tryb dzienny' : 'Tryb nocny';
  try { localStorage.theme = dark ? 'dark' : 'light'; } catch(e){}
}

// ========== Data loading ==========
async function loadDataStructure() {
  try {
    const res = await fetch('/api/get_data_structure');
    if (!res.ok) throw new Error('Błąd sieci: ' + res.status);
    const data = await res.json();
    if (!refs.pion) return;
    refs.pion.innerHTML = '';
    refs.pion.append(new Option('-- Wybierz --',''));
    Object.keys(data || {}).forEach(k => refs.pion.append(new Option(k,k)));
  } catch (e) {
    console.error('Błąd ładowania struktury danych:', e);
    showMessage('Błąd ładowania danych. Sprawdź serwer (konsola).', true);
  }
}

async function loadGT() {
  resetFormExceptPion();
  const p = refs.pion ? refs.pion.value : '';
  const uniqInfo = document.getElementById('uniqueFormatInfo');
  if (uniqInfo) uniqInfo.style.display = p === 'Oświetlenie' ? 'block' : 'none';
  if (refs.gtInput) refs.gtInput.disabled = p === 'Oświetlenie';
  if (p === 'Oświetlenie') { disableKW(); return; }
  if (!p) return;
  try {
    const res = await fetch(`/api/get_gt?pion=${encodeURIComponent(p)}`);
    if (!res.ok) throw new Error('Błąd get_gt: ' + res.status);
    availableGTs = await res.json();
    refreshGTDataList();
  } catch (e) {
    console.error('Błąd pobierania GT:', e);
    showMessage('Błąd pobierania GT', true);
  }
}

function refreshGTDataList() {
  if (!refs.gtList) return;
  const dl = refs.gtList;
  const sel = refs.gtTags ? [...refs.gtTags.children].map(ch => ch.dataset.value) : [];
  dl.innerHTML = '';
  availableGTs.filter(gt => !sel.includes(gt)).forEach(gt => dl.append(new Option(gt,gt)));
}

function updateGTList() {
  if (!refs.gtInput || !refs.gtList) return;
  const inp = refs.gtInput.value.trim().toLowerCase();
  const sel = refs.gtTags ? [...refs.gtTags.children].map(ch => ch.dataset.value.toLowerCase()) : [];
  refs.gtList.innerHTML = '';
  availableGTs
    .filter(gt => !sel.includes(gt.toLowerCase()) && (!inp || gt.toLowerCase().includes(inp)))
    .forEach(gt => refs.gtList.append(new Option(gt,gt)));
}

function selectGTTag() {
  if (!refs.gtInput || !refs.gtTags) return;
  const v = refs.gtInput.value.trim();
  if (!v) return;
  if (![...refs.gtTags.children].some(ch => ch.dataset.value === v)) {
    const sp = document.createElement('span');
    sp.className='tag'; sp.dataset.value=v; sp.textContent=v;
    const btn = document.createElement('button'); btn.type='button'; btn.textContent='✕';
    sp.appendChild(btn);
    refs.gtTags.appendChild(sp);
    updateKW(); refreshGTDataList();
  }
  refs.gtInput.value = '';
}

function disableKW() {
  if (refs.kwInput) refs.kwInput.disabled = true;
  if (refs.kwList) refs.kwList.innerHTML = '';
  if (refs.kwTags) refs.kwTags.innerHTML = '';
}

async function updateKW() {
  const gtList = refs.gtTags ? [...refs.gtTags.children].map(ch => ch.dataset.value) : [];
  if (!gtList.length) return disableKW();
  try {
    const res = await fetch('/api/get_kw_for_gt_list', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ gtList })
    });
    if (!res.ok) throw new Error('Błąd get_kw: ' + res.status);
    const arr = await res.json();
    availableKWs = arr;
    if (!refs.kwList) return;
    refs.kwList.innerHTML = '';
    arr.forEach(k => refs.kwList.append(new Option(k,k)));
    refs.kwInput.disabled = false;
  } catch (e) {
    console.error('Błąd pobierania KW:', e);
    showMessage('Błąd pobierania KW', true);
  }
}

function selectKWTag() {
  if (!refs.kwInput || !refs.kwTags) return;
  const v = refs.kwInput.value.trim();
  if (!v) return;
  if (![...refs.kwTags.children].some(ch => ch.dataset.value === v)) {
    const sp = document.createElement('span');
    sp.className='tag'; sp.dataset.value=v; sp.textContent=v;
    const btn = document.createElement('button'); btn.type='button'; btn.textContent='✕';
    sp.appendChild(btn);
    refs.kwTags.appendChild(sp);
  }
  refs.kwInput.value = '';
}

// ========== Paste handlers ==========
refs.gtInput?.addEventListener('paste', async function(e) {
  try {
    e.preventDefault();
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!text) return;
    const res = await fetch('/api/resolve_gt_codes', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ pion: refs.pion ? refs.pion.value : '', raw: text })
    });
    if (!res.ok) throw new Error('Błąd resolve_gt_codes: ' + res.status);
    const fullGTs = await res.json();
    if (!Array.isArray(fullGTs)) return;
    fullGTs.forEach(gtName => {
      refs.gtInput.value = gtName;
      selectGTTag();
    });
  } catch (err) {
    console.error('resolve paste error', err);
    showMessage('Nie udało się rozszerzyć kodów GT', true);
  }
});

refs.kwInput?.addEventListener('paste', function(e) {
  try {
    e.preventDefault();
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!text) return;
    text.split(',').map(s => s.trim()).filter(s => s && availableKWs.includes(s)).forEach(val => {
      refs.kwInput.value = val;
      selectKWTag();
    });
    refs.kwInput.value = '';
  } catch (err) {
    console.error('KW paste error', err);
  }
});

// remove tag handler
document.addEventListener('click', function(e) {
  if (e.target && e.target.matches && e.target.matches('.tags button')) {
    const sp = e.target.closest('span.tag');
    if (!sp) return;
    const parent = sp.parentElement;
    sp.remove();
    if (parent && parent.id === 'gtTags') { refreshGTDataList(); updateKW(); }
  }
});

// ========== Submit ==========
function validateEmailList(raw) {
  if (!raw) return [];
  const parts = raw.split(/[;,]/).map(s => s.trim()).filter(Boolean);
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return parts.filter(p => re.test(p));
}

async function submitForm(e) {
  e.preventDefault();
  clearMessage();
  const pion = refs.pion ? refs.pion.value : '';
  const gtList = refs.gtTags ? [...refs.gtTags.children].map(ch => ch.dataset.value) : [];
  const kwList = refs.kwTags ? [...refs.kwTags.children].map(ch => ch.dataset.value) : [];
  const emailRaw = refs.email ? refs.email.value.trim() : '';
  const emails = validateEmailList(emailRaw);

  if (!pion) { showMessage('Wybierz pion', true); return; }
  if (pion !== 'Oświetlenie' && (gtList.length === 0 || kwList.length === 0)) { showMessage('Wybierz GT i KW', true); return; }
  if (!emails.length) { showMessage('Podaj poprawny adres e-mail', true); return; }

  setLoading(true);
  saveLastEmail(emailRaw);
  // remove previous download link if any
  const dlWrap = document.getElementById('downloadWrapper'); if (dlWrap) dlWrap.remove();

  try {
    const res = await fetch('/api/generate', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ pion, gtList, kwList, email: emailRaw })
    });
    const data = await res.json();
    if (res.ok && data.success) {
      // show nice message
      showMessage(`✅ Plik został wygenerowany i wysłany. Sprawdź skrzynkę mailową.<br><br>Nazwa pliku: <strong>${data.filename || 'formatki.xlsx'}</strong>`);
      // jeśli backend zwróci link do pliku (file_url lub url), pokaż przycisk
      const fileUrl = data.file_url || data.url || data.fileUrl || data.download_url;
      if (fileUrl) createDownloadButton(fileUrl, data.filename || 'Pobierz plik');
    } else {
      console.error('generate error:', data);
      showMessage(`❌ Błąd generowania: ${data.error || 'Nieznany błąd'}`, true);
    }
  } catch (err) {
    console.error('Submit error', err);
    showMessage('Błąd połączenia z serwerem podczas wysyłki', true);
  } finally {
    setLoading(false);
  }
}

// ========== Init ==========
document.addEventListener('DOMContentLoaded', function() {
  try {
    // theme
    if (localStorage.theme === 'dark') {
      document.body.classList.add('dark-mode');
      if (refs.themeIcon) refs.themeIcon.textContent = '☀️';
      if (refs.themeText) refs.themeText.textContent = 'Tryb dzienny';
    }
    // load saved email
    loadLastEmail();

    loadDataStructure();
    if (refs.pion) refs.pion.addEventListener('change', loadGT);
    if (refs.gtInput) {
      refs.gtInput.addEventListener('input', updateGTList);
      refs.gtInput.addEventListener('keydown', e => e.key === 'Enter' && (e.preventDefault(), selectGTTag()));
      refs.gtInput.addEventListener('change', selectGTTag);
    }
    if (refs.kwInput) {
      refs.kwInput.addEventListener('keydown', e => e.key === 'Enter' && (e.preventDefault(), selectKWTag()));
      refs.kwInput.addEventListener('change', selectKWTag);
    }
    if (refs.themeToggle) refs.themeToggle.addEventListener('click', toggleTheme);
    const form = document.getElementById('myForm');
    if (form) form.addEventListener('submit', submitForm);
    console.log('app.js (blok 3) załadowany');
  } catch (err) {
    console.error('Init error:', err);
  }
});
