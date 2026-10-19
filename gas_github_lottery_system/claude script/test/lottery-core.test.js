// P0 單元測試：node --test "claude script/test/"
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const C = require('../project/lottery-core.js');
const XlsxLite = require('../project/xlsx-lite.js');

const FIXTURE = path.join(__dirname, 'fixtures', '測試用抽獎後台.xlsx');

async function loadFixture() {
  const buf = fs.readFileSync(FIXTURE);
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  return C.fromWorkbook(await XlsxLite.readXlsx(ab));
}
async function fixtureState() {
  const got = await loadFixture();
  return { ...C.createState(), source: 'workbook', participants: got.participants, prizes: got.prizes };
}
const prizeByName = (s, kw) => s.prizes.find(p => p.name.includes(kw));

/* ---------- 匯入（G1、G2） ---------- */
test('自動辨識測試 Excel 的參加者與獎品工作表', async () => {
  const got = await loadFixture();
  assert.equal(got.participants.length, 18);
  assert.equal(got.prizes.length, 4);
  assert.equal(got.peopleSheet, '表單回覆 1');
  assert.equal(got.prizeSheet, '獎品設定');
  assert.deepEqual(got.prizes.map(p => p.quota), [10, 2, 5, 10]);
  assert.ok(got.prizes[0].name.startsWith('頭獎'), '依抽獎順序排序');
});

test('參加者保留系所、年級、學號，且學號不帶小數點', async () => {
  const got = await loadFixture();
  const p = got.participants.find(x => x.name === '張育誠');
  assert.equal(p.studentId, '413777367');
  assert.equal(p.id, 'sid-413777367');
  assert.equal(p.department, 'AI');
  assert.equal(p.grade, '大四');
  assert.equal(p.email, undefined, '不載入電子郵件');
});

test('同名不同學號視為兩個人；學號重複時改用列號並警告（B4）', () => {
  const rows = [['姓名', '學號'], ['王小明', '1'], ['王小明', '2'], ['陳大文', '1'], ['林美麗', '']];
  const { participants, warnings } = C.participantsFromRows(rows);
  assert.deepEqual(participants.map(p => p.id), ['sid-1', 'sid-2', 'row-4', 'row-5']);
  assert.equal(warnings.length, 1);
});

test('手動名單：同名者有不同 ID', () => {
  const ps = C.participantsFromLines('王小明\n\n王小明\n陳大文\n');
  assert.deepEqual(ps.map(p => p.id), ['row-1', 'row-2', 'row-3']);
});

test('沒有獎項工作表時使用預設獎項，名額 = 人數', () => {
  const got = C.fromWorkbook({ SheetNames: ['a'], data: { a: [['姓名'], ['甲'], ['乙']] } });
  assert.equal(got.prizes.length, 1);
  assert.equal(got.prizes[0].quota, 2);
});

test('不是名單的工作表回傳 null（交給手動選欄位）', () => {
  assert.equal(C.fromWorkbook({ SheetNames: ['a'], data: { a: [['x', 'y'], ['1', '2']] } }), null);
});

/* ---------- 名額與資格 ---------- */
test('抽獎扣名額，且中獎者不會再出現在任何獎項', async () => {
  const s = await fixtureState();
  const second = prizeByName(s, '二獎');
  const made = C.draw(s, second.id, 1);
  assert.equal(made.length, 1);
  assert.equal(C.remainingQuota(s, second.id), 1);
  for (const p of s.prizes) assert.ok(!C.eligible(s, p.id).some(x => x.id === made[0].participantId));
});

test('要求人數超過剩餘名額時只抽到名額上限；名額 0 時抽不到', async () => {
  const s = await fixtureState();
  const second = prizeByName(s, '二獎');
  assert.equal(C.draw(s, second.id, 5).length, 2);
  assert.equal(C.remainingQuota(s, second.id), 0);
  assert.equal(C.draw(s, second.id, 1).length, 0);
});

test('要求人數超過合格人數時只抽到合格人數', () => {
  const s = { ...C.createState(), participants: C.participantsFromLines('甲\n乙'),
              prizes: [{ id: 'p', name: 'P', quota: 10, order: 1 }] };
  assert.equal(C.draw(s, 'p', 10).length, 2);
  assert.equal(C.maxDrawable(s, 'p'), 0);
});

test('同一次抽獎不會抽到重複的人（B5）', async () => {
  const s = await fixtureState();
  for (const p of s.prizes) C.draw(s, p.id, p.quota);
  const ids = s.records.map(r => r.participantId);
  assert.equal(new Set(ids).size, ids.length);
  assert.equal(ids.length, 18, '總名額 27 > 18 人，全部人都抽到後停止');
});

test('commitDraw 會擋下不合格或超額的寫入（防止畫面與狀態不同步，B1）', async () => {
  const s = await fixtureState();
  const second = prizeByName(s, '二獎');
  const picks = C.pick(s, second.id, 1);
  C.commitDraw(s, second.id, picks);
  assert.throws(() => C.commitDraw(s, second.id, picks), /合格名單/);
  const two = C.pick(s, prizeByName(s, '頭獎').id, 2);
  assert.throws(() => C.commitDraw(s, second.id, two), /剩餘名額/);
});

test('抽獎結果是均勻隨機（簡單統計檢查）', () => {
  const counts = { 甲: 0, 乙: 0, 丙: 0 };
  for (let i = 0; i < 3000; i++) {
    const s = { ...C.createState(), participants: C.participantsFromLines('甲\n乙\n丙'),
                prizes: [{ id: 'p', name: 'P', quota: 1, order: 1 }] };
    counts[C.draw(s, 'p', 1)[0].who.name]++;
  }
  for (const k in counts) assert.ok(counts[k] > 850 && counts[k] < 1150, JSON.stringify(counts));
});

/* ---------- 時間、匯出、保存 ---------- */
test('時間包含完整日期（B6）', () => {
  assert.equal(C.formatTime(new Date(2026, 9, 19, 9, 5, 3)), '2026-10-19 09:05:03');
});

test('CSV 含獎項、系所、學號與完整時間，並正確跳脫逗號與引號', () => {
  const s = { ...C.createState(), participants: [{ id: 'a', name: '王"小明', department: 'AI,四', grade: '大四', studentId: '1' }],
              prizes: [{ id: 'p', name: '頭獎', quota: 1, order: 1 }] };
  C.draw(s, 'p', 1, { now: new Date(2026, 9, 19, 9, 5, 3) });
  const csv = C.toCSV(s);
  assert.ok(csv.startsWith('﻿序號,獎項,姓名,系所,年級,學號,抽出時間'));
  assert.ok(csv.includes('1,頭獎,"王""小明","AI,四",大四,1,2026-10-19 09:05:03'));
});

test('保存後還原：獎項、名單、紀錄都保留；pending 不保存', async () => {
  const s = await fixtureState();
  C.draw(s, s.prizes[0].id, 3);
  s.records.push({ ...s.records[0], id: 'x', status: 'pending' });
  const back = C.deserialize(C.serialize(s));
  assert.equal(back.records.length, 3);
  assert.equal(back.participants.length, 18);
  assert.equal(C.remainingQuota(back, s.prizes[0].id), 7);
});

test('損壞或舊版的保存資料會被忽略', () => {
  assert.equal(C.deserialize('{壞掉'), null);
  assert.equal(C.deserialize(JSON.stringify({ version: 1 })), null);
  assert.equal(C.deserialize(null), null);
});

test('清除紀錄後名額恢復', async () => {
  const s = await fixtureState();
  C.draw(s, s.prizes[0].id, 4);
  C.resetRecords(s);
  assert.equal(C.remainingQuota(s, s.prizes[0].id), 10);
  assert.equal(C.eligible(s, s.prizes[0].id).length, 18);
});
