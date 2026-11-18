import {reportRequests, reportTables, workPhotoTypes} from './google-report.js?v=1';

const driveScope = 'https://www.googleapis.com/auth/drive.file';
const drive = 'https://www.googleapis.com/drive/v3';
const sheets = 'https://sheets.googleapis.com/v4/spreadsheets';
const folderType = 'application/vnd.google-apps.folder';
const sheetType = 'application/vnd.google-apps.spreadsheet';
const clientIdKey = 'tattoo-crm-google-client-id';
const stateKey = 'tattoo-crm-google-sync-v1';
const clientFields = ['name', 'phone', 'telegram', 'birthday', 'comment'];
const headings = ['ID клиента', 'Имя', 'Телефон', 'Telegram', 'Дата рождения', 'Комментарий', 'ID фотографии', 'SHA-256', 'Удалён'];
const safeId = /^[a-zA-Z0-9_-]{1,128}$/;
const safeClientId = /^[a-zA-Z0-9_-]{1,48}$/;
const hexDigest = /^[a-f0-9]{64}$/;
const maxPhotoSize = 15 * 1024 * 1024;

class GoogleError extends Error {
  constructor(status, url, responseError) {
    super('Google API: ' + status);
    this.status = status;
    const {hostname, pathname} = new URL(url);
    this.service = hostname === 'sheets.googleapis.com' ? 'Google Sheets API'
      : pathname.includes('/drive/v3/') ? 'Google Drive API' : 'Вход через Google';
    this.reasons = new Set([
      ...(responseError?.details || []).map(detail => detail.reason),
      ...(responseError?.errors || []).map(detail => detail.reason)
    ].filter(reason => typeof reason === 'string' && /^[a-zA-Z0-9_]{1,80}$/.test(reason)));
  }
}

function serviceUrl(base, path, params = {}) {
  const url = new URL(base + path);
  for (const [key, value] of Object.entries(params)) url.searchParams.set(key, value);
  return url.href;
}

function loadState() {
  try {
    return JSON.parse(localStorage.getItem(stateKey)) || {};
  } catch {
    return {};
  }
}

function saveState(patch) {
  localStorage.setItem(stateKey, JSON.stringify({...loadState(), ...patch}));
}

function clientValues(client) {
  return clientFields.map(field => String(client[field] ?? ''));
}

function fingerprint(record) {
  return record ? JSON.stringify([clientValues(record.client), record.photoHash]) : '';
}

async function photoHash(blob) {
  if (!blob) return '';
  const bytes = await crypto.subtle.digest('SHA-256', await blob.arrayBuffer());
  return Array.from(new Uint8Array(bytes), byte => byte.toString(16).padStart(2, '0')).join('');
}

function googleStorage(session) {
  async function request(url, {method = 'GET', json, body, type, binary} = {}) {
    const headers = {Authorization: 'Bearer ' + session.token};
    if (json !== undefined) {
      headers['Content-Type'] = 'application/json';
      body = JSON.stringify(json);
    } else if (type) {
      headers['Content-Type'] = type;
    }
    const response = await fetch(url, {
      method, headers, body, signal: AbortSignal.timeout(45000), cache: 'no-store'
    });
    if (!response.ok) {
      const data = await response.json().catch(() => null);
      throw new GoogleError(response.status, url, data?.error);
    }
    if (binary) return response.blob();
    if (response.status !== 204) return response.json();
  }

  async function findFile(role, parent, extra = '') {
    const query = ["trashed = false", "appProperties has { key='tattooRole' and value='" + role + "' }"];
    if (parent) query.push("'" + parent + "' in parents");
    if (extra) query.push(extra);
    let pageToken;
    do {
      const result = await request(serviceUrl(drive, '/files', {
        q: query.join(' and '), fields: 'files(id),nextPageToken', orderBy: 'createdTime', pageSize: 10,
        ...(pageToken ? {pageToken} : {})
      }));
      if (result.files?.length) return result.files[0].id;
      pageToken = result.nextPageToken;
    } while (pageToken);
  }

  async function findOrCreate(role, name, mimeType, parent) {
    const existing = await findFile(role, parent);
    if (existing) return existing;
    const file = {name, mimeType, appProperties: {tattooRole: role}};
    if (parent) file.parents = [parent];
    const result = await request(serviceUrl(drive, '/files', {fields: 'id'}), {method: 'POST', json: file});
    return result.id;
  }

  function valueUrl(workbook, range, append = false) {
    const path = '/' + workbook.id + '/values/' + encodeURIComponent("'" + workbook.title.replaceAll("'", "''") + "'!" + range);
    return serviceUrl(sheets, path + (append ? ':append' : ''), {
      valueInputOption: 'RAW', ...(append ? {insertDataOption: 'INSERT_ROWS'} : {})
    });
  }

  async function openWorkbook() {
    if (session.workbook) return session.workbook;
    const root = await findOrCreate('root-v1', 'Tattoo CRM', folderType);
    const photos = await findOrCreate('photos-v1', 'Фото', folderType, root);
    const id = await findOrCreate('clients-v1', 'Клиенты Tattoo CRM', sheetType, root);
    const info = await request(serviceUrl(sheets, '/' + id, {fields: 'sheets(properties(sheetId,title))'}));
    const worksheet = info.sheets?.find(item => item.properties.title === 'Клиенты') || info.sheets?.[0];
    if (!worksheet) throw new Error('В Google Таблице нет листа для клиентов.');
    const {sheetId, title} = worksheet.properties;
    const workbook = {id, title, photos, root};
    const {values} = await request(serviceUrl(sheets, '/' + id + '/values/' + encodeURIComponent("'" + title.replaceAll("'", "''") + "'!A1:I1")));
    if (!values?.length) {
      if (title !== 'Клиенты') {
        await request(sheets + '/' + id + ':batchUpdate', {
          method: 'POST',
          json: {requests: [{updateSheetProperties: {properties: {sheetId, title: 'Клиенты'}, fields: 'title'}}]}
        });
        workbook.title = 'Клиенты';
      }
      await request(valueUrl(workbook, 'A1:I1'), {method: 'PUT', json: {values: [headings]}});
    } else if (JSON.stringify(values[0]) !== JSON.stringify(headings)) {
      throw new Error('Не удалось распознать колонки листа «Клиенты». Проверь первую строку таблицы.');
    }
    session.workbook = workbook;
    return workbook;
  }

  async function listClients() {
    const workbook = await openWorkbook();
    const range = "'" + workbook.title.replaceAll("'", "''") + "'!A2:I";
    const {values = []} = await request(serviceUrl(sheets, '/' + workbook.id + '/values/' + encodeURIComponent(range), {
      valueRenderOption: 'UNFORMATTED_VALUE'
    }));
    const remote = new Map();
    values.forEach((cells, index) => {
      if (!cells.some(cell => cell !== '')) return;
      const [id, name = '', phone = '', telegram = '', birthday = '', comment = '', photoId = '', hash = '', removed = ''] = cells.map(String);
      if (!safeClientId.test(id || '') || remote.has(id) || (removed !== '1' && (!name.trim() ||
          Boolean(photoId) !== Boolean(hash) || (photoId && (!safeId.test(photoId) || !hexDigest.test(hash)))))) {
        throw new Error('Проверь строку ' + (index + 2) + ' листа «Клиенты». CRM не смогла прочитать её.');
      }
      remote.set(id, {
        row: index + 2,
        record: removed === '1' ? null : {client: {id, name, phone, telegram, birthday, comment}, photoId, photoHash: hash}
      });
    });
    return remote;
  }

  async function writeClient(id, record, remote) {
    const workbook = await openWorkbook();
    const row = remote?.row;
    const values = [record
      ? [id, ...clientValues(record.client), record.photoId, record.photoHash, '']
      : [id, '', '', '', '', '', '', '', '1']];
    const range = row ? 'A' + row + ':I' + row : 'A:I';
    if (row) {
      const path = '/' + workbook.id + '/values/' + encodeURIComponent("'" + workbook.title.replaceAll("'", "''") + "'!" + range);
      const {values: current} = await request(serviceUrl(sheets, path, {valueRenderOption: 'UNFORMATTED_VALUE'}));
      const old = remote.record;
      const expected = old ? [id, ...clientValues(old.client), old.photoId, old.photoHash, ''] : [id, '', '', '', '', '', '', '', '1'];
      const actual = Array.from({length: headings.length}, (_, index) => String(current?.[0]?.[index] ?? ''));
      if (JSON.stringify(actual) !== JSON.stringify(expected)) {
        throw new Error('Строка в Google изменилась во время синхронизации. CRM проверит её повторно.');
      }
    }
    await request(valueUrl(workbook, range, !row), {method: row ? 'PUT' : 'POST', json: {values}});
  }

  async function putPhoto(id, hash, blob) {
    const {photos} = await openWorkbook();
    return putImage(photos, 'client-photo-v1', id, hash, blob, id);
  }

  async function putWorkPhoto(id, hash, blob, name) {
    const {photos} = await openWorkbook();
    session.workFolder ||= await findOrCreate('work-photos-v1', 'Эскизы и фото работ', folderType, photos);
    return putImage(session.workFolder, 'work-photo-v1', id, hash, blob, name);
  }

  async function putImage(parent, role, id, hash, blob, name) {
    const existing = await findFile(role, parent, "appProperties has { key='tattooPhoto' and value='" + id + ':' + hash + "' }");
    if (existing) return existing;
    const boundary = 'tattoo-' + crypto.randomUUID();
    const extension = {'image/jpeg': '.jpg', 'image/png': '.png', 'image/webp': '.webp', 'image/heic': '.heic', 'image/gif': '.gif'}[blob.type] || '';
    const metadata = {
      name: name + extension, mimeType: blob.type, parents: [parent],
      appProperties: {tattooRole: role, tattooPhoto: id + ':' + hash}
    };
    const body = new Blob([
      '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n',
      JSON.stringify(metadata), '\r\n--' + boundary + '\r\nContent-Type: ' + blob.type + '\r\n\r\n',
      blob, '\r\n--' + boundary + '--\r\n'
    ]);
    const response = await request(serviceUrl('https://www.googleapis.com/upload/drive/v3', '/files', {
      uploadType: 'multipart', fields: 'id'
    }), {method: 'POST', body, type: 'multipart/related; boundary=' + boundary});
    return response.id;
  }

  async function writeReport(deviceId, tables, previousId, unchanged) {
    const {root} = await openWorkbook();
    const id = await findOrCreate('report-' + deviceId, 'Отчёт Tattoo CRM · ' + deviceId.slice(0, 8), sheetType, root);
    if (unchanged && id === previousId) return id;
    const info = await request(serviceUrl(sheets, '/' + id, {fields: 'sheets(properties(sheetId,title))'}));
    const requests = reportRequests(tables, (info.sheets || []).map(sheet => sheet.properties));
    await request(sheets + '/' + id + ':batchUpdate', {
      method: 'POST',
      json: {requests: [{updateSpreadsheetProperties: {properties: {locale: 'ru_RU'}, fields: 'locale'}}, ...requests]}
    });
    return id;
  }

  async function getPhoto(id, expectedHash) {
    const blob = await request(serviceUrl(drive, '/files/' + id, {alt: 'media'}), {binary: true});
    if (!blob.type.startsWith('image/') || blob.size > maxPhotoSize || await photoHash(blob) !== expectedHash) {
      throw new Error('Не удалось проверить фотографию на Google Диске.');
    }
    return blob;
  }

  async function trashPhoto(fileId, photoKey) {
    const {photos} = await openWorkbook();
    const path = '/files/' + fileId;
    let file;
    try {
      file = await request(serviceUrl(drive, path, {fields: 'parents,appProperties,trashed'}));
    } catch (error) {
      if (error.status === 404) return;
      throw error;
    }
    if (file.trashed || !file.parents?.includes(photos) || file.appProperties?.tattooPhoto !== photoKey) return;
    await request(serviceUrl(drive, path, {fields: 'trashed'}), {method: 'PATCH', json: {trashed: true}});
  }

  return {profile: () => request('https://openidconnect.googleapis.com/v1/userinfo'), listClients, writeClient, putPhoto, getPhoto, trashPhoto, putWorkPhoto, writeReport};
}

export function mountGoogleSync(panel, local) {
  const element = id => panel.querySelector('#' + id);
  const settings = element('googleSetup');
  const clientIdInput = element('googleClientId');
  const connectButton = element('googleConnect');
  const transferButton = element('googleTransfer');
  const disconnectButton = element('googleDisconnect');
  const keepButton = element('googleKeepLocal');
  const restoreButton = element('googleKeepRemote');
  const workbookLink = element('googleWorkbook');
  const reportLink = element('googleReportLink');
  const reportToggle = element('googleReportToggle');
  const status = element('googleStatus');
  const savedClientId = localStorage.getItem(clientIdKey) || '';
  const initialState = loadState();
  const pairedClientId = initialState.account ? initialState.oauthClientId || savedClientId : '';
  const sharedClientId = panel.dataset.googleClientId.trim();
  const useSharedClient = Boolean(sharedClientId && (!pairedClientId || pairedClientId === sharedClientId));
  const configuredClientId = useSharedClient ? sharedClientId : pairedClientId || savedClientId;
  let oauthClient;
  let session;
  let loading = false;
  let pendingAuth = false;
  let working = false;
  let retryAfterRun = false;
  let conflicts = 0;
  let timer;
  let expiresTimer;
  let localVersion = 0;
  const tell = message => { status.textContent = message; };

  function refreshControls() {
    const state = loadState();
    connectButton.disabled = loading || pendingAuth || working || !oauthClient;
    connectButton.classList.toggle('hidden', Boolean(session));
    transferButton.classList.toggle('hidden', !session);
    transferButton.disabled = working;
    transferButton.textContent = state.enabled ? 'Синхронизировать сейчас' : 'Начать синхронизацию';
    disconnectButton.classList.toggle('hidden', !session && !state.enabled);
    disconnectButton.disabled = pendingAuth || working;
    settings.querySelector('button').disabled = loading || pendingAuth || working;
    for (const button of [keepButton, restoreButton]) {
      button.classList.toggle('hidden', !session || !conflicts);
      button.disabled = working;
    }
    workbookLink.classList.toggle('hidden', !session?.workbook);
    if (session?.workbook) workbookLink.href = 'https://docs.google.com/spreadsheets/d/' + session.workbook.id + '/edit';
    reportToggle.checked = Boolean(state.reportEnabled);
    reportToggle.disabled = working;
    reportLink.classList.toggle('hidden', !state.reportId);
    if (safeId.test(state.reportId || '')) reportLink.href = 'https://docs.google.com/spreadsheets/d/' + state.reportId + '/edit';
  }

  function clearSession() {
    clearTimeout(expiresTimer);
    session = undefined;
    conflicts = 0;
    refreshControls();
  }

  function reportError(error) {
    if (error.status === 401) {
      clearSession();
      tell('Доступ к Google истёк. Войди через Google, чтобы продолжить синхронизацию.');
    } else if (error.status === 403 && (error.reasons.has('SERVICE_DISABLED') || error.reasons.has('accessNotConfigured'))) {
      tell(error.service + ' отключён. Включи этот API в том проекте Google Cloud, где создан Client ID CRM.');
    } else if (error.status === 403 && (error.reasons.has('ACCESS_TOKEN_SCOPE_INSUFFICIENT') || error.reasons.has('insufficientPermissions'))) {
      clearSession();
      tell('Google не выдал доступ к файлам CRM. Войди снова и разреши доступ к файлам приложения.');
    } else if (error.status === 403) {
      const reason = [...error.reasons].join(', ');
      tell(error.service + ' отклонил запрос (403' + (reason ? ', ' + reason : '') + ').');
    } else if (error.status === 429 || error.status >= 500 || error.name === 'TimeoutError' || error instanceof TypeError) {
      tell('Google пока недоступен. Изменения остались на устройстве, CRM повторит попытку.');
    } else {
      tell(error instanceof GoogleError ? 'Не удалось синхронизировать с Google (' + error.status + ').' : error.message);
    }
  }

  async function acceptToken(response) {
    pendingAuth = false;
    if (response.error || !response.access_token || !google.accounts.oauth2.hasGrantedAllScopes(response, driveScope)) {
      tell('Для синхронизации нужен доступ к файлам Tattoo CRM на Google Диске.');
      refreshControls();
      return;
    }
    working = true;
    refreshControls();
    try {
      const candidate = {token: response.access_token};
      const profile = await googleStorage(candidate).profile();
      if (!profile.sub) throw new Error('Google не вернул идентификатор аккаунта.');
      const state = loadState();
      if (state.account && state.account !== profile.sub) throw new Error('Это устройство связано с другим аккаунтом Google. Войди в ранее подключённый аккаунт.');
      if (useSharedClient) saveState({enabled: true, account: profile.sub, oauthClientId: sharedClientId});
      session = {...candidate, account: profile.sub, email: profile.email || ''};
      expiresTimer = setTimeout(() => {
        clearSession();
        tell('Доступ к Google истёк. Изменения сохраняются на устройстве до подключения.');
      }, Math.max(0, (Number(response.expires_in) || 3600) * 1000 - 10000));
      tell('Google подключён' + (session.email ? ' · ' + session.email : '') + '.');
    } catch (error) {
      reportError(error);
    } finally {
      working = false;
      refreshControls();
    }
    if (!session) return;
    if (loadState().enabled) synchronize();
  }

  async function setup(clientId) {
    oauthClient = undefined;
    if (!clientId) {
      tell('Вход через Google ещё не настроен. Для активации нужен Client ID приложения.');
      refreshControls();
      return;
    }
    loading = true;
    refreshControls();
    try {
      if (!window.google?.accounts?.oauth2) {
        await new Promise((resolve, reject) => {
          const script = document.createElement('script');
          script.src = 'https://accounts.google.com/gsi/client';
          script.async = true;
          script.onload = resolve;
          script.onerror = () => { script.remove(); reject(new Error('Не удалось загрузить вход Google. Проверь интернет.')); };
          document.head.append(script);
        });
      }
      oauthClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: 'openid email ' + driveScope,
        include_granted_scopes: false,
        callback: acceptToken,
        error_callback: () => {
          pendingAuth = false;
          tell('Окно входа Google закрыто или заблокировано. Попробуй подключиться ещё раз.');
          refreshControls();
        }
      });
      tell(loadState().enabled ? 'Войди через Google, чтобы продолжить синхронизацию.' : 'Синхронизация с Google выключена.');
    } catch (error) {
      tell(error.message);
    } finally {
      loading = false;
      refreshControls();
    }
  }

  function notePhoto(id) {
    const state = loadState();
    if (state.account) saveState({dirty: {...state.dirty, [id]: crypto.randomUUID()}});
    schedule();
  }

  function schedule() {
    localVersion++;
    clearTimeout(timer);
    if (!loadState().enabled) return;
    if (!session) {
      tell('Изменения сохранены на устройстве. Войди через Google для синхронизации.');
      return;
    }
    if (working) retryAfterRun = true;
    timer = setTimeout(() => synchronize(), 900);
  }

  async function syncClients(active, resolution) {
    const state = loadState();
    const storage = googleStorage(active);
    const known = new Map(Object.entries(state.known || {}));
    const version = localVersion;
    const localClients = new Map(local.clients().filter(client => !client.demo && safeClientId.test(client.id || '')).map(client => [client.id, client]));
    const remoteClients = await storage.listClients();
    const ids = new Set([...localClients.keys(), ...known.keys(), ...remoteClients.keys()]);
    conflicts = 0;
    let deferred = false;

    for (const id of ids) {
      if (session !== active || localVersion !== version) return;
      const prior = known.get(id);
      const current = localClients.get(id);
      const remote = remoteClients.get(id);
      const cloud = remote?.record || null;
      const dirty = Object.hasOwn(state.dirty || {}, id) ? state.dirty[id] : undefined;
      let photo;
      if (current && (!prior || dirty)) photo = await local.getPhoto(id);
      const currentHash = current ? (!prior || dirty ? await photoHash(photo) : prior.photoHash) : '';
      const device = current ? {client: current, photoHash: currentHash, photoId: prior?.photoId || ''} : null;
      const deviceStamp = fingerprint(device);
      const cloudStamp = fingerprint(cloud);
      const baseStamp = prior?.stamp || '';
      const localChanged = deviceStamp !== baseStamp;
      const remoteChanged = cloudStamp !== baseStamp;
      const diverged = localChanged && remoteChanged && deviceStamp !== cloudStamp;
      if (diverged && !resolution) {
        conflicts++;
        continue;
      }
      if (session !== active || localVersion !== version) return;
      let synced = cloud;
      if (deviceStamp !== cloudStamp && (diverged ? resolution === 'local' : localChanged)) {
        if (device?.photoHash) {
          if (cloud?.photoHash === device.photoHash) device.photoId = cloud.photoId;
          else {
            photo ||= await local.getPhoto(id);
            if (!photo || !photo.type.startsWith('image/') || photo.size > maxPhotoSize ||
                await photoHash(photo) !== device.photoHash) throw new Error('Фото изменилось во время синхронизации. Попробуй ещё раз.');
            device.photoId = await storage.putPhoto(id, device.photoHash, photo);
          }
        } else if (device) device.photoId = '';
        if (session !== active || localVersion !== version) return;
        if (cloud?.photoId && cloud.photoId !== device?.photoId) {
          saveState({retiredPhotos: {...loadState().retiredPhotos, [cloud.photoId]: id + ':' + cloud.photoHash}});
        }
        await storage.writeClient(id, device, remote);
        synced = device;
      } else if (deviceStamp !== cloudStamp) {
        if (!cloud && local.hasSessions(id)) {
          conflicts++;
          continue;
        }
        if (local.isEditing()) {
          deferred = true;
          continue;
        }
        const incomingPhoto = cloud?.photoHash === currentHash
          ? undefined : cloud?.photoId ? await storage.getPhoto(cloud.photoId, cloud.photoHash) : null;
        if (session !== active || localVersion !== version || local.isEditing()) return;
        await local.receive(id, cloud?.client, incomingPhoto);
      }
      const stamp = fingerprint(synced);
      if (stamp !== baseStamp || prior?.photoId !== synced?.photoId || dirty) {
        if (synced) known.set(id, {stamp, photoId: synced.photoId, photoHash: synced.photoHash});
        else known.delete(id);
        const latest = loadState().dirty || {};
        if (dirty && latest[id] === dirty) delete latest[id];
        saveState({known: Object.fromEntries(known), dirty: latest});
      }
    }
    const retired = Object.entries(loadState().retiredPhotos || {});
    if (retired.length && session === active) {
      const latest = await storage.listClients();
      const referenced = new Set([...latest.values()].map(row => row.record?.photoId));
      for (const [id, key] of retired) {
        if (session !== active) return;
        if (safeId.test(id) && !referenced.has(id)) await storage.trashPhoto(id, key);
        const remaining = loadState().retiredPhotos || {};
        delete remaining[id];
        saveState({retiredPhotos: remaining});
      }
    }
    if (conflicts) return 'Есть несовпадения в карточках (' + conflicts + '). Выбери, какую версию сохранить.';
    if (deferred) return 'Загрузка изменений продолжится после закрытия карточки.';
  }

  async function syncReport(active, version) {
    const state = loadState();
    if (!state.reportEnabled) return;
    const device = state.reportDevice || crypto.randomUUID();
    if (!state.reportDevice) saveState({reportDevice: device});
    const storage = googleStorage(active);
    const dataset = local.reportData();
    const keys = new Set(await local.workPhotoKeys());
    const saved = state.workPhotos || {};
    const uploaded = {};
    for (const appointment of dataset.appointments) {
      if (!safeClientId.test(appointment.id || '')) continue;
      for (const [prefix, label] of workPhotoTypes) {
        const key = prefix + appointment.id;
        if (!keys.has(key)) continue;
        if (session !== active || localVersion !== version) return;
        const marker = state.workChanges?.[key];
        const existing = saved[key];
        if (existing?.marker === marker && existing?.fileId) {
          uploaded[key] = existing;
          continue;
        }
        const blob = await local.workPhoto(key);
        if (!blob || !blob.type.startsWith('image/') || blob.size > maxPhotoSize) continue;
        const hash = await photoHash(blob);
        if (session !== active || localVersion !== version) return;
        const name = [appointment.date, label, appointment.id.slice(0, 8), hash.slice(0, 8)].join(' — ');
        const fileId = await storage.putWorkPhoto(key, hash, blob, name);
        uploaded[key] = {fileId, marker};
        saveState({workPhotos: {...loadState().workPhotos, [key]: uploaded[key]}});
      }
    }
    if (session !== active || localVersion !== version) return;
    dataset.appointments = dataset.appointments.filter(appointment =>
      !appointment.demo || workPhotoTypes.some(([prefix]) => keys.has(prefix + appointment.id)));
    const tables = reportTables(dataset, uploaded);
    const hash = await photoHash(new Blob([JSON.stringify(tables)]));
    if (session !== active || localVersion !== version) return;
    const id = await storage.writeReport(device, tables, state.reportId, state.reportHash === hash);
    saveState({reportId: id, reportHash: hash, workPhotos: uploaded});
  }

  async function syncAll(active, resolution) {
    const version = localVersion;
    const message = await syncClients(active, resolution);
    if (session !== active || localVersion !== version) return;
    await syncReport(active, version);
    if (session === active && localVersion === version) {
      tell(message || 'Синхронизировано ' + new Date().toLocaleTimeString('ru-RU', {hour: '2-digit', minute: '2-digit'}) + '.');
    }
  }

  async function synchronize(resolution) {
    if (working || !session || !loadState().enabled) return;
    if (!navigator.onLine) {
      tell('Нет интернета. Изменения сохранены на устройстве.');
      return;
    }
    working = true;
    const active = session;
    tell('Синхронизация с Google…');
    refreshControls();
    try {
      if (navigator.locks) {
        const ran = await navigator.locks.request('tattoo-crm-google', {ifAvailable: true}, async lock => {
          if (!lock) return false;
          await syncAll(active, resolution);
          return true;
        });
        if (!ran) tell('Синхронизация уже идёт в другой вкладке.');
      } else await syncAll(active, resolution);
    } catch (error) {
      reportError(error);
    } finally {
      working = false;
      refreshControls();
      if (retryAfterRun) {
        retryAfterRun = false;
        clearTimeout(timer);
        timer = setTimeout(() => synchronize(), 900);
      }
    }
  }

  settings.addEventListener('submit', event => {
    event.preventDefault();
    const clientId = clientIdInput.value.trim();
    if (!/^[a-zA-Z0-9_-]+\.apps\.googleusercontent\.com$/.test(clientId)) {
      tell('Укажи OAuth Client ID веб-приложения из Google Cloud.');
      return;
    }
    const state = loadState();
    const previous = state.oauthClientId || localStorage.getItem(clientIdKey);
    if (state.account && previous !== clientId) {
      tell('На этом устройстве уже настроена синхронизация с другим OAuth Client ID.');
      return;
    }
    localStorage.setItem(clientIdKey, clientId);
    clearSession();
    setup(clientId);
  });
  connectButton.addEventListener('click', () => {
    if (!oauthClient || pendingAuth) return;
    pendingAuth = true;
    tell('Ожидаем вход в Google…');
    refreshControls();
    try {
      oauthClient.requestAccessToken({prompt: loadState().account ? '' : 'select_account'});
    } catch {
      pendingAuth = false;
      tell('Не удалось открыть вход в Google. Попробуй снова.');
      refreshControls();
    }
  });
  transferButton.addEventListener('click', () => {
    if (!session) return;
    if (!loadState().enabled) {
      if (!confirm('Подключить синхронизацию карточек и фотографий клиентов с этим аккаунтом Google? Существующие данные могут быть отправлены на Google Диск.')) return;
      saveState({enabled: true, account: session.account, oauthClientId: localStorage.getItem(clientIdKey)});
    }
    synchronize();
  });
  disconnectButton.addEventListener('click', () => {
    clearSession();
    saveState({enabled: false});
    tell('Синхронизация приостановлена. Данные остаются на устройстве и в Google.');
    refreshControls();
  });
  keepButton.addEventListener('click', () => {
    if (confirm('Отправить версии карточек с этого устройства вместо конфликтующих версий в Google?')) synchronize('local');
  });
  restoreButton.addEventListener('click', () => {
    if (confirm('Заменить конфликтующие версии карточек на этом устройстве версиями из Google?')) synchronize('remote');
  });
  reportToggle.addEventListener('change', () => {
    saveState({reportEnabled: reportToggle.checked});
    schedule();
  });
  const trySync = () => { if (!document.hidden) synchronize(); };
  window.addEventListener('online', trySync);
  document.addEventListener('visibilitychange', trySync);
  setInterval(trySync, 60000);
  const setupDetails = settings.closest('details');
  setupDetails.hidden = useSharedClient;
  clientIdInput.value = pairedClientId || savedClientId;
  if (configuredClientId) setupDetails.open = false;
  element('googleSyncDisclosure').classList.toggle('hidden', !useSharedClient);
  setup(configuredClientId);
  return {changed: schedule, photoChanged: notePhoto, workPhotoChanged(key) {
    const state = loadState();
    if (state.reportDevice) saveState({workChanges: {...state.workChanges, [key]: crypto.randomUUID()}});
    schedule();
  }};
}
