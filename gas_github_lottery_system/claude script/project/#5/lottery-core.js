/* 抽獎系統 2.0 — 核心邏輯（不碰 DOM，可在瀏覽器與 Node 執行）
 * 瀏覽器：window.LotteryCore；Node：module.exports
 *
 * P0：獎項 / 名額、參加者穩定 ID、中獎紀錄、保存與還原、CSV
 * #6：抽出後為 'pending'（待確認），確認後才成為 'confirmed'；捨棄則刪除該筆
 * #5：正式中獎可取消為 'cancelled'（記錄原因），名額退回一次；可針對取消紀錄重抽一位補位
 */
(function (root) {
  'use strict';

  const STORAGE_KEY = 'lottery_v2_state';
  const VERSION = 2;
  const PERSISTED_STATUSES = ['confirmed', 'cancelled']; // pending 不保存（Q2）

  /* ---------- 亂數（沿用上週的無偏差加密亂數，可注入以便測試） ---------- */
  function cryptoRandInt(max) {
    const limit = Math.floor(0x100000000 / max) * max, a = new Uint32Array(1);
    do globalThis.crypto.getRandomValues(a); while (a[0] >= limit);
    return a[0] % max;
  }
  function shuffle(arr, randInt = cryptoRandInt) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) { const j = randInt(i + 1); [a[i], a[j]] = [a[j], a[i]]; }
    return a;
  }

  /* ---------- 工具 ---------- */
  const pad = n => String(n).padStart(2, '0');
  // 修正 B6：完整日期時間，跨日領獎也能對帳
  function formatTime(d) {
    d = d instanceof Date ? d : new Date(d);
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ` +
           `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
  }
  const clean = v => String(v ?? '').trim();
  // Excel 數字欄位可能變成 "413777367.0" 或科學記號
  function cleanStudentId(v) {
    let s = clean(v);
    if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
    if (/^\d+(\.\d+)?e\+?\d+$/i.test(s)) s = String(Math.round(Number(s)));
    return s;
  }
  function findCol(head, aliases) {
    const h = head.map(x => clean(x));
    for (const a of aliases) { const i = h.findIndex(x => x === a); if (i >= 0) return i; }
    for (const a of aliases) { const i = h.findIndex(x => x.includes(a)); if (i >= 0) return i; }
    return -1;
  }

  /* ---------- 狀態 ---------- */
  function createState() {
    return { version: VERSION, source: 'none', prizes: [], participants: [], records: [] };
  }

  /* ---------- 匯入：參加者 ---------- */
  // 修正 B4 / G2：每位參加者有穩定 ID（學號，沒有學號就用列號），並保留系所、年級
  function participantsFromRows(rows) {
    const warnings = [];
    if (!rows.length) return { participants: [], warnings: ['參加者工作表是空的'] };
    const head = rows[0];
    const ci = {
      name: findCol(head, ['姓名', '名字', 'name']),
      dept: findCol(head, ['系所', '科系', '系級', 'department']),
      grade: findCol(head, ['年級', 'grade']),
      sid: findCol(head, ['學號', 'student id', 'id']),
    };
    if (ci.name < 0) return null; // 不是參加者工作表
    const participants = [], seen = new Set();
    rows.slice(1).forEach((r, i) => {
      const rowNo = i + 2; // Excel 列號（含標題列）
      const name = clean(r[ci.name]);
      if (!name) { if (r.some(x => clean(x))) warnings.push(`第 ${rowNo} 列沒有姓名，已略過`); return; }
      const sid = ci.sid >= 0 ? cleanStudentId(r[ci.sid]) : '';
      let id = sid ? `sid-${sid}` : `row-${rowNo}`;
      if (seen.has(id)) {
        warnings.push(`第 ${rowNo} 列學號 ${sid} 與前面重複，暫以列號區分`);
        id = `row-${rowNo}`;
      }
      seen.add(id);
      participants.push({
        id, name, studentId: sid,
        department: ci.dept >= 0 ? clean(r[ci.dept]) : '',
        grade: ci.grade >= 0 ? clean(r[ci.grade]) : '',
      });
    });
    // 電子郵件刻意不載入（Q9：不收集不必要的個資）
    return { participants, warnings };
  }

  // 手動貼上名單：每行一人，ID 用列號（同名者也會是不同的人）
  function participantsFromLines(text) {
    return String(text).split(/\r?\n/).map(clean).filter(Boolean)
      .map((name, i) => ({ id: `row-${i + 1}`, name, studentId: '', department: '', grade: '' }));
  }

  /* ---------- 匯入：獎項 ---------- */
  function prizesFromRows(rows) {
    const warnings = [];
    if (!rows.length) return null;
    const head = rows[0];
    const ci = {
      name: findCol(head, ['獎品名稱', '獎項名稱', '獎品', '獎項', 'prize']),
      quota: findCol(head, ['配額', '名額', '數量', 'quota']),
      order: findCol(head, ['抽獎順序', '順序', 'order']),
      note: findCol(head, ['備註', '說明', 'description']),
    };
    if (ci.name < 0 || ci.quota < 0) return null; // 不是獎項工作表
    const prizes = [];
    rows.slice(1).forEach((r, i) => {
      const rowNo = i + 2;
      const name = clean(r[ci.name]);
      if (!name) return;
      const quota = Number(clean(r[ci.quota]));
      if (!Number.isInteger(quota) || quota < 0) {
        warnings.push(`獎項「${name}」的配額不是有效的整數，已略過`); return;
      }
      const order = ci.order >= 0 && clean(r[ci.order]) !== '' ? Number(clean(r[ci.order])) : rowNo;
      prizes.push({ id: `prize-${rowNo}`, name, quota, order: Number.isFinite(order) ? order : rowNo,
                    note: ci.note >= 0 ? clean(r[ci.note]) : '' });
    });
    prizes.sort((a, b) => a.order - b.order);
    return { prizes, warnings };
  }

  // 沒有獎項工作表時的預設獎項：名額 = 參加人數（等同上週「抽過不放回」的行為）
  function defaultPrize(count) {
    return { id: 'prize-default', name: '抽籤', quota: count, order: 1, note: '' };
  }

  /* 自動辨識活頁簿：找出參加者工作表與獎項工作表
   * book = { SheetNames, data: {name: rows} }（XlsxLite.readXlsx 的輸出） */
  function fromWorkbook(book) {
    let people = null, prizes = null, peopleSheet = '', prizeSheet = '';
    for (const name of book.SheetNames) {
      const rows = book.data[name] || [];
      if (!prizes) { const p = prizesFromRows(rows); if (p) { prizes = p; prizeSheet = name; continue; } }
      if (!people) { const q = participantsFromRows(rows); if (q) { people = q; peopleSheet = name; } }
    }
    if (!people) return null;
    const warnings = [...people.warnings, ...(prizes ? prizes.warnings : [])];
    const prizeList = prizes && prizes.prizes.length ? prizes.prizes : [defaultPrize(people.participants.length)];
    if (!prizes) warnings.push('找不到獎項工作表，改用單一預設獎項');
    return { participants: people.participants, prizes: prizeList, peopleSheet, prizeSheet, warnings };
  }

  /* ---------- 名額與資格 ---------- */
  const byId = (list, id) => list.find(x => x.id === id);
  const countsTowardQuota = r => r.status === 'confirmed' || r.status === 'pending';

  function usedQuota(state, prizeId) {
    return state.records.filter(r => r.prizeId === prizeId && countsTowardQuota(r)).length;
  }
  function remainingQuota(state, prizeId) {
    const p = byId(state.prizes, prizeId);
    return p ? Math.max(0, p.quota - usedQuota(state, prizeId)) : 0;
  }
  // 合格名單：排除任何獎項的正式 / 待確認得主、以及在此獎項被取消過的人（Q3）
  function eligible(state, prizeId) {
    const out = new Set();
    for (const r of state.records) {
      if (countsTowardQuota(r)) out.add(r.participantId);
      else if (r.status === 'cancelled' && r.prizeId === prizeId) out.add(r.participantId);
    }
    return state.participants.filter(p => !out.has(p.id));
  }
  function maxDrawable(state, prizeId) {
    return Math.min(remainingQuota(state, prizeId), eligible(state, prizeId).length);
  }

  /* ---------- 抽獎 ---------- */
  // 只挑人，不寫入紀錄（讓畫面可以先播動畫，再呼叫 commitDraw）
  function pick(state, prizeId, n, randInt = cryptoRandInt) {
    const count = Math.min(Math.max(0, Math.floor(n) || 0), maxDrawable(state, prizeId));
    if (count < 1) return [];
    return shuffle(eligible(state, prizeId), randInt).slice(0, count);
  }

  let seq = 0;
  const newId = () => `r-${Date.now().toString(36)}-${(seq++).toString(36)}`;

  // 寫入紀錄前再檢查一次名額與資格，避免畫面與狀態不同步時寫入錯誤資料
  // #6：預設寫入為 pending，名額先預扣，但不進正式名單與 CSV
  function commitDraw(state, prizeId, people, now = new Date(), status = 'pending', extra = {}) {
    if (!byId(state.prizes, prizeId)) throw new Error('找不到這個獎項');
    const ok = new Set(eligible(state, prizeId).map(p => p.id));
    if (people.length > remainingQuota(state, prizeId)) throw new Error('超過此獎項剩餘名額');
    if (new Set(people.map(p => p.id)).size !== people.length) throw new Error('同一次抽出重複的人');
    for (const p of people) if (!ok.has(p.id)) throw new Error(`${p.name} 已不在此獎項的合格名單中`);
    const t = formatTime(now);
    const made = people.map(p => ({
      id: newId(), prizeId, participantId: p.id, status,
      who: { name: p.name, department: p.department, grade: p.grade, studentId: p.studentId },
      drawnAt: t, confirmedAt: status === 'confirmed' ? t : null,
      isRedraw: !!extra.replacesId, replacesId: extra.replacesId || null,
    }));
    state.records.push(...made);
    return made;
  }

  function draw(state, prizeId, n, opts = {}) {
    return commitDraw(state, prizeId, pick(state, prizeId, n, opts.randInt), opts.now);
  }

  function resetRecords(state) { state.records = []; }

  /* ---------- #6 確認 / 捨棄 ---------- */
  const pendingRecords = state => state.records.filter(r => r.status === 'pending');
  const hasPending = state => state.records.some(r => r.status === 'pending');

  function confirm(state, id, now = new Date()) {
    const r = byId(state.records, id);
    if (!r) throw new Error('找不到這筆紀錄');
    if (r.status !== 'pending') throw new Error('只有待確認的結果可以確認');
    r.status = 'confirmed';
    r.confirmedAt = formatTime(now);
    return r;
  }
  function confirmAll(state, now = new Date()) {
    return pendingRecords(state).map(r => confirm(state, r.id, now));
  }
  // 捨棄：直接刪除（它從來不是正式結果），名額與資格自然恢復（Q1）
  function discard(state, id) {
    const r = byId(state.records, id);
    if (!r) throw new Error('找不到這筆紀錄');
    if (r.status !== 'pending') throw new Error('只有待確認的結果可以捨棄');
    state.records = state.records.filter(x => x.id !== id);
    return r;
  }
  function discardAll(state) {
    const gone = pendingRecords(state);
    state.records = state.records.filter(r => r.status !== 'pending');
    return gone;
  }

  /* ---------- #5 取消中獎與重抽 ---------- */
  const CANCEL_REASONS = ['不在場', '資格不符', '其他'];

  // 只能取消「正式中獎」且「未領獎」的紀錄（Q7、Q10）；已取消的不能再取消 → 名額只退一次
  function canCancel(r) {
    if (!r) return '找不到這筆紀錄';
    if (r.status === 'cancelled') return '這筆紀錄已經取消過了';
    if (r.status !== 'confirmed') return '只有正式中獎的紀錄可以取消';
    if (r.claimed) return '已領獎，請先改回未領獎才能取消';
    return '';
  }
  function cancel(state, id, reason, note = '', now = new Date()) {
    const r = byId(state.records, id);
    const why = canCancel(r);
    if (why) throw new Error(why);
    if (!CANCEL_REASONS.includes(reason)) throw new Error('請選擇取消原因');
    r.status = 'cancelled';
    r.cancelReason = reason;
    r.cancelNote = reason === '其他' ? clean(note) : '';
    r.cancelledAt = formatTime(now);
    return r;
  }

  // 找出補這筆取消紀錄的重抽紀錄。被捨棄的 pending 已刪除，所以可以再重抽；
  // 若補位者後來也被取消，就改由「那一筆」繼續重抽，避免同一個名額出現兩個重抽按鈕
  const replacementOf = (state, cancelledId) =>
    state.records.find(r => r.replacesId === cancelledId) || null;

  function canRedraw(state, cancelledId) {
    const c = byId(state.records, cancelledId);
    if (!c || c.status !== 'cancelled') return '只有已取消的紀錄可以重抽';
    if (replacementOf(state, cancelledId)) return '這筆已經有人補位了';
    if (hasPending(state)) return '請先確認或捨棄待確認的結果';
    if (remainingQuota(state, c.prizeId) < 1) return '此獎項名額已被用完，無法重抽';
    if (!eligible(state, c.prizeId).length) return '此獎項已沒有可抽的人';
    return '';
  }
  // 與一般抽獎相同：先挑人（播動畫），再 commitRedraw 寫入為 pending（Q6）
  function pickRedraw(state, cancelledId, randInt = cryptoRandInt) {
    const why = canRedraw(state, cancelledId);
    if (why) throw new Error(why);
    const c = byId(state.records, cancelledId);
    return pick(state, c.prizeId, 1, randInt);
  }
  function commitRedraw(state, cancelledId, people, now = new Date()) {
    const why = canRedraw(state, cancelledId);
    if (why) throw new Error(why);
    const c = byId(state.records, cancelledId);
    return commitDraw(state, c.prizeId, people.slice(0, 1), now, 'pending', { replacesId: cancelledId })[0];
  }

  /* ---------- 匯出 ---------- */
  const csvCell = v => {
    const s = String(v ?? '');
    return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const STATUS_LABEL = { confirmed: '正式', cancelled: '已取消' };
  // Q11：匯出正式與已取消的紀錄（待確認不匯出）
  function toCSV(state) {
    const head = ['序號', '獎項', '姓名', '系所', '年級', '學號', '狀態', '是否重抽', '補位對象',
                  '抽出時間', '確認時間', '取消原因', '取消時間'];
    const rows = state.records.filter(r => r.status === 'confirmed' || r.status === 'cancelled').map((r, i) => {
      const prize = byId(state.prizes, r.prizeId);
      const replaced = r.replacesId ? byId(state.records, r.replacesId) : null;
      const reason = r.cancelReason ? r.cancelReason + (r.cancelNote ? `：${r.cancelNote}` : '') : '';
      return [i + 1, prize ? prize.name : r.prizeId, r.who.name, r.who.department, r.who.grade,
              r.who.studentId, STATUS_LABEL[r.status], r.isRedraw ? '是' : '否',
              replaced ? replaced.who.name : '', r.drawnAt, r.confirmedAt || '', reason, r.cancelledAt || ''];
    });
    return '﻿' + [head, ...rows].map(r => r.map(csvCell).join(',')).join('\r\n');
  }

  /* ---------- 保存 / 還原（G3） ---------- */
  function serialize(state) {
    return JSON.stringify({
      version: VERSION, source: state.source, prizes: state.prizes, participants: state.participants,
      records: state.records.filter(r => PERSISTED_STATUSES.includes(r.status)),
    });
  }
  function deserialize(json) {
    if (!json) return null;
    let o;
    try { o = JSON.parse(json); } catch { return null; }
    if (!o || o.version !== VERSION || !Array.isArray(o.prizes) ||
        !Array.isArray(o.participants) || !Array.isArray(o.records)) return null;
    return { version: VERSION, source: o.source || 'none', prizes: o.prizes, participants: o.participants,
             records: o.records.filter(r => PERSISTED_STATUSES.includes(r.status)) };
  }

  const api = {
    STORAGE_KEY, VERSION, formatTime, cleanStudentId, shuffle, cryptoRandInt,
    createState, participantsFromRows, participantsFromLines, prizesFromRows, defaultPrize, fromWorkbook,
    usedQuota, remainingQuota, eligible, maxDrawable, pick, commitDraw, draw, resetRecords,
    pendingRecords, hasPending, confirm, confirmAll, discard, discardAll,
    CANCEL_REASONS, canCancel, cancel, replacementOf, canRedraw, pickRedraw, commitRedraw,
    toCSV, serialize, deserialize,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LotteryCore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
