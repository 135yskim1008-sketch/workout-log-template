// 운동 기록 앱 ↔ 이 구글 시트 자동 연동
// 기록은 이 시트의 '기록' 탭에만 저장됩니다. 아래 비밀번호는 앱이 만들어 준 것이니 바꾸지 마세요.
const SECRET = "%KEY%";
const HEAD = ["날짜", "운동", "부위", "무게kg", "세트", "횟수", "메모", "id", "수정시각"];
const CM = "코멘트·방향";

function doGet(e) {
  return out({ ok: true, msg: "운동 기록 연동이 켜져 있어요." });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || "{}");
    if (body.key !== SECRET) return out({ ok: false, error: "비밀번호가 달라요" });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("기록") || ss.insertSheet("기록");
    const ts = ss.getSheetByName("삭제기록") || ss.insertSheet("삭제기록");
    const tz = ss.getSpreadsheetTimeZone();

    // 1) 시트에 있는 기록 읽기
    const rows = {};
    const n = sh.getLastRow() - 1;
    const vals = n > 0 ? sh.getRange(2, 1, n, 9).getValues() : [];
    vals.forEach(function (v) {
      const id = String(v[7] || "");
      if (!id) return;
      rows[id] = {
        date: v[0] instanceof Date ? Utilities.formatDate(v[0], tz, "yyyy-MM-dd") : String(v[0]),
        ex: String(v[1]), part: String(v[2]), kg: String(v[3]), set: String(v[4]), rep: String(v[5]),
        memo: String(v[6]), id: id, u: Number(v[8]) || 0
      };
    });
    const tomb = {};
    const tn = ts.getLastRow() - 1;
    const tv = tn > 0 ? ts.getRange(2, 1, tn, 2).getValues() : [];
    tv.forEach(function (v) { if (v[0]) tomb[String(v[0])] = Number(v[1]) || 0; });

    // 2) 앱에서 온 기록 합치기 (나중에 고친 것이 이김, 지운 기록은 빼기)
    const bt = body.tomb || {};
    Object.keys(bt).forEach(function (k) { tomb[k] = Math.max(tomb[k] || 0, Number(bt[k]) || 0); });
    (body.rows || []).forEach(function (r) {
      if (!r.id) return;
      const u = Number(r.u) || 0;
      const o = rows[r.id];
      if (!o || u >= o.u) {
        rows[r.id] = {
          date: String(r.date), ex: String(r.ex), part: String(r.part || ""),
          kg: r.kg == null ? "" : String(r.kg), set: r.set == null ? "" : String(r.set),
          rep: r.rep == null ? "" : String(r.rep), memo: String(r.memo || ""), id: String(r.id), u: u
        };
      }
    });
    const list = Object.keys(rows).map(function (k) { return rows[k]; })
      .filter(function (r) { return !(tomb[r.id] && tomb[r.id] >= r.u); });
    list.sort(function (a, b) {
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      return (a.ex === CM ? 1 : 0) - (b.ex === CM ? 1 : 0);
    });

    // 3) 시트에 다시 쓰기
    sh.clear();
    const data = [HEAD].concat(list.map(function (r) { return [r.date, r.ex, r.part, r.kg, r.set, r.rep, r.memo, r.id, String(r.u)]; }));
    sh.getRange(1, 1, data.length, 9).setNumberFormat("@").setValues(data);
    sh.setFrozenRows(1);
    ts.clear();
    const td = [["id", "삭제시각"]].concat(Object.keys(tomb).map(function (k) { return [k, String(tomb[k])]; }));
    ts.getRange(1, 1, td.length, 2).setNumberFormat("@").setValues(td);

    return out({ ok: true, rows: list, tomb: tomb });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
