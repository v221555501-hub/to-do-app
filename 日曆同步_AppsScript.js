/**
 * 待辦 App → Google 日曆「自動同步」（Apps Script 網頁應用程式）
 * App 新增／修改／刪除／完成任務，或按「📅 加入 Google 日曆」時，會自動在你的 Google 日曆
 * 建立、更新或刪除事件，手機的 Google 日曆就會依提醒時間跳通知。
 *
 * 事件：開始日期時間 → 結束日期時間（沒填時間：開始預設 09:00、結束預設 18:00，台北時間）
 * 提醒：開始時間當下，以及「提前通知」的時間（天／小時／分鐘，以開始時間為基準）
 * 顏色：紅→紅、黃→黃、綠→綠、藍→藍、紫→紫、灰→灰
 * 已完成或取消的任務，會自動從日曆刪除。
 *
 * 另外：App 按「匯入／更新 Google 日曆」時，會讀取你的 Google 帳號「看得到」的所有日曆
 * （自己的、別人共用給你並已訂閱的），把事件傳回 App 顯示。對方沒共用的日曆讀不到。
 */
const SECRET = '請改成你自己的密語';   // 必須與 App「⚙ 設定」裡的同步密語一致
const TZ = 'Asia/Taipei';

function doGet() { return ContentService.createTextOutput('OK'); }

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.tok !== SECRET) return ContentService.createTextOutput('denied');
  if (b.a === 'all') (b.tasks || []).forEach(upsert_);
  else if (b.a === 'up') upsert_(b.task);
  else if (b.a === 'del') remove_(b.task.id);
  else if (b.a === 'events') return json_(events_(b.from, b.to));
  return ContentService.createTextOutput('ok');
}

function colorOf_(c) {
  const E = CalendarApp.EventColor;
  return { red: E.RED, yellow: E.YELLOW, green: E.GREEN, blue: E.BLUE, purple: E.MAUVE, gray: E.GRAY }[c] || E.GRAY;
}

function upsert_(t) {
  if (!t || !t.id) return;
  if (t.prog === 1 || t.prog === '取消') return remove_(t.id);        // 完成／取消 → 刪除事件
  const start = Utilities.parseDate(t.start + ' ' + (t.sT || '09:00'), TZ, 'yyyy-MM-dd HH:mm');
  let end = Utilities.parseDate(t.due + ' ' + (t.eT || '18:00'), TZ, 'yyyy-MM-dd HH:mm');
  if (end <= start) end = new Date(start.getTime() + 3600000);
  const title = t.title;
  const desc = [t.folder && '資料夾：' + t.folder, t.cat && '分類：' + t.cat,
                t.contact && '主要聯繫人：' + t.contact,
                '重要程度：' + ['☆', '☆☆', '★★★', '★★★★'][t.imp],
                t.note && '追蹤進度／內容：\n' + t.note].filter(Boolean).join('\n');
  const props = PropertiesService.getScriptProperties(), key = 't_' + t.id;
  const cal = CalendarApp.getDefaultCalendar();
  let ev = null;
  try { const id = props.getProperty(key); if (id) ev = cal.getEventById(id); } catch (err) {}
  if (ev) { ev.setTitle(title); ev.setTime(start, end); ev.setDescription(desc); }
  else { ev = cal.createEvent(title, start, end, { description: desc }); props.setProperty(key, ev.getId()); }
  ev.setColor(colorOf_(t.col));
  ev.removeAllReminders();
  ev.addPopupReminder(0);                                              // 開始時間當下
  const mult = { '天': 1440, '小時': 60, '分鐘': 1 }[t.advU] || 1440;
  const mins = Math.round((Number(t.advN) || 0) * mult);
  if (mins > 0 && mins <= 40320) ev.addPopupReminder(mins);            // 提前通知（Google 日曆上限 4 週）
}

function remove_(id) {
  const props = PropertiesService.getScriptProperties(), key = 't_' + id, evId = props.getProperty(key);
  if (!evId) return;
  try { const ev = CalendarApp.getDefaultCalendar().getEventById(evId); if (ev) ev.deleteEvent(); } catch (err) {}
  props.deleteProperty(key);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 讀取帳號下所有日曆（含訂閱的他人日曆）在 from～to（yyyy-MM-dd）之間的事件，排除本 App 自己建立的事件。 */
function events_(from, to) {
  const start = Utilities.parseDate(from + ' 00:00', TZ, 'yyyy-MM-dd HH:mm');
  const end = Utilities.parseDate(to + ' 23:59', TZ, 'yyyy-MM-dd HH:mm');
  const mine = new Set(Object.values(PropertiesService.getScriptProperties().getProperties()));
  const list = [], out = [];
  CalendarApp.getAllCalendars().forEach(c => {
    list.push({ id: c.getId(), name: c.getName(), color: c.getColor() });
    c.getEvents(start, end).forEach(ev => {
      if (out.length >= 1500 || mine.has(ev.getId())) return;
      const allDay = ev.isAllDayEvent();
      const s = ev.getStartTime();
      const e = allDay ? new Date(ev.getEndTime().getTime() - 86400000) : ev.getEndTime();
      out.push({
        id: c.getId() + '|' + ev.getId(), cal: c.getId(), cn: c.getName(), col: c.getColor(),
        title: ev.getTitle(),
        start: Utilities.formatDate(s, TZ, 'yyyy-MM-dd'), sT: allDay ? '' : Utilities.formatDate(s, TZ, 'HH:mm'),
        due: Utilities.formatDate(e, TZ, 'yyyy-MM-dd'), eT: allDay ? '' : Utilities.formatDate(e, TZ, 'HH:mm'),
        loc: ev.getLocation(), note: String(ev.getDescription() || '').slice(0, 500)
      });
    });
  });
  return { cals: list, events: out };
}
