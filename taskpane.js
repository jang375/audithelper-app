/* global Office, Excel */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    loadSheets();
    const bind = (id, fn) => { const el = document.getElementById(id); if (el) el.onclick = fn; };
    document.getElementById('refresh-sheets').onclick = loadSheets;

    bind('s1-src-range-btn', () => getSelectionRange('s1-src-range'));
    bind('s1-tgt-range-btn', () => getSelectionRange('s1-tgt-range'));
    bind('s2-src-range-btn', () => getSelectionRange('s2-src-range'));
    bind('s2-tgt-range-btn', () => getSelectionRange('s2-tgt-range'));

    bind('s1-diag', () => runStep1('DIAG'));
    bind('s1-sync-cur', () => runStep1('SYNC_CUR'));
    bind('s1-sync-pri', () => runStep1('SYNC_PRI'));
    bind('s1-verify', () => runStep1Verify());

    bind('s2-diag', () => runStep2('DIAG'));
    bind('s2-sync-cur', () => runStep2('SYNC_CUR'));
    bind('s2-sync-pri', () => runStep2('SYNC_PRI'));
    bind('s2-verify', () => runStep2Verify());

    window.goToCell = goToCell;
  }
});

/* ──────── 🚀 v8.8 초정밀 엔진 (셀 참조 전송 지원) ──────── */

const contraKeywords = ["대손충당금", "감가상각누계액", "현재가치할인차금", "정부보조금", "보조금", "손상차손누계액", "손상차손"];
const forceDataItemKeywords = ["이익잉여금", "당기순이익", "당기순손실"]; 

const superClean = (val) => {
  if (!val) return "";
  let s = String(val).replace(/[\s\uFEFF\xA0\u200B\u200C\u200D\u202F\u205F\u3000]+/g, "").trim();
  s = s.replace(/^[0-9]+[\.\s]*/, "").replace(/^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVX]+[\.\s]*/i, "").replace(/^\([0-9]+\)/, "").replace(/^[가-힣][\.]/, "");
  return s.replace(/[^가-힣a-zA-Z0-9]/g, "");
};

const isRoman = (t) => /^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVX]+\./i.test(String(t).trim());
const isSubHeader = (t) => /^\([0-9]+\)/.test(String(t).trim());
const isGrandTotal = (t) => { const cn = superClean(t); return cn.includes("총계") || cn.includes("합계"); };

const isSumOrHeader = (name) => {
  if (forceDataItemKeywords.some(k => superClean(name).includes(k))) return false;
  return isRoman(name) || isSubHeader(name) || isGrandTotal(name);
};

const getColLetter = (n) => {
  let letter = "";
  while (n >= 0) { letter = String.fromCharCode((n % 26) + 65) + letter; n = Math.floor(n / 26) - 1; }
  return letter;
};

const parseNum = (val) => {
  const n = Number(String(val ?? "0").replace(/[^0-9.-]/g, ""));
  return isNaN(n) ? 0 : n;
};

/* ──────── 시스템 유틸리티 ──────── */

async function getSelectionRange(inputId) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      document.getElementById(inputId).value = range.address.split('!')[1] || range.address;
    });
  } catch (e) { console.error(e); }
}

async function loadSheets() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets.load('items/name');
    await context.sync();
    const selects = ['s1-src-sheet', 's1-tgt-sheet', 's2-src-sheet', 's2-tgt-sheet'];
    selects.forEach(id => {
      const el = document.getElementById(id);
      if(el) { el.innerHTML = ''; sheets.items.forEach(s => el.add(new Option(s.name, s.name))); }
    });
  });
}

async function goToCell(sheetName, absRow) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.getRangeByIndexes(absRow, 0, 1, 1).select();
    sheet.activate();
    await context.sync();
  });
}

/* ──────── [1단계] 동기화 로직 (유지) ──────── */

async function runStep1(mode) {
  const summary = document.getElementById('summary-text'); const list = document.getElementById('mismatch-list');
  const cfg = {
    srcS: document.getElementById('s1-src-sheet').value, srcR: document.getElementById('s1-src-range').value,
    tgtS: document.getElementById('s1-tgt-sheet').value, tgtR: document.getElementById('s1-tgt-range').value,
    srcAcc: parseInt(document.getElementById('s1-src-acc').value)-1, srcCur: parseInt(document.getElementById('s1-src-cur').value)-1, srcPri: parseInt(document.getElementById('s1-src-pri').value)-1,
    tgtAcc: parseInt(document.getElementById('s1-tgt-acc').value)-1, tgtCur: parseInt(document.getElementById('s1-tgt-cur').value)-1, tgtPri: parseInt(document.getElementById('s1-tgt-pri').value)-1
  };
  try {
    summary.innerText = "분석 중..."; list.innerHTML = "";
    await Excel.run(async (context) => {
      const sS = context.workbook.worksheets.getItem(cfg.srcS); const tS = context.workbook.worksheets.getItem(cfg.tgtS);
      const sRO = sS.getRange(cfg.srcR).load(["values", "rowIndex"]);
      const tRO = tS.getRange(cfg.tgtR).load(["values", "rowIndex", "columnIndex"]);
      await context.sync();
      const tSC = tRO.columnIndex;
      const tgtStrictMap = new Map(), tgtFuzzyMap = new Map();
      let lastTgtMain = "Root";
      tRO.values.forEach((row, idx) => {
        const raw = row[cfg.tgtAcc]; if(!raw) return;
        const cn = superClean(raw);
        const isH = isSumOrHeader(raw);
        const effectiveParent = !isH && contraKeywords.some(k => cn.includes(k)) ? lastTgtMain : "Root";
        const key = `${effectiveParent}_${cn}`;
        const data = { absRow: tRO.rowIndex + idx, name: raw, cn: cn };
        if(!tgtStrictMap.has(key)) tgtStrictMap.set(key, []);
        tgtStrictMap.get(key).push(data);
        if(!tgtFuzzyMap.has(cn)) tgtFuzzyMap.set(cn, []);
        tgtFuzzyMap.get(cn).push(data);
        if(!isH) lastTgtMain = cn;
      });
      const pending = [], results = [], tgtUse = {}, fuzzyUse = {}, contraUse = {};
      let srcLastMain = "Root";
      sRO.values.forEach((row, idx) => {
        const raw = row[cfg.srcAcc]; if(!raw) return;
        const cn = superClean(raw);
        const isH = isSumOrHeader(raw);
        const effectiveParent = !isH && contraKeywords.some(k => cn.includes(k)) ? srcLastMain : "Root";
        const key = `${effectiveParent}_${cn}`;
        let m = null;
        if(tgtStrictMap.has(key)) { const matches = tgtStrictMap.get(key); const use = tgtUse[key] || 0; if(matches[use]) { m = matches[use]; tgtUse[key] = use + 1; } }
        if(!m && effectiveParent !== "Root") {
            const parentItems = []; tgtStrictMap.forEach((val, k) => { if(k.startsWith(effectiveParent + "_")) parentItems.push(...val); });
            const uKey = `CONTRA_${effectiveParent}_${cn}`; const use = contraUse[uKey] || 0;
            const candidates = parentItems.filter(item => item.cn.includes(cn) || cn.includes(item.cn));
            if(candidates[use]) { m = candidates[use]; contraUse[uKey] = use + 1; }
        }
        if(!m && tgtFuzzyMap.has(cn)) { const inst = tgtFuzzyMap.get(cn); const use = fuzzyUse[cn] || 0; if(inst[use]) { m = inst[use]; fuzzyUse[cn] = use + 1; } }
        const cV = parseNum(row[cfg.srcCur]), pV = parseNum(row[cfg.srcPri]);
        if(m) {
          if(mode==='SYNC_CUR') pending.push({r: m.absRow, c: tSC + cfg.tgtCur, v: cV});
          if(mode==='SYNC_PRI') pending.push({r: m.absRow, c: tSC + cfg.tgtPri, v: pV});
        } else if(cV !== 0 || pV !== 0) { results.push({name: raw, absRow: sRO.rowIndex + idx}); }
        if(!isH) srcLastMain = cn;
      });
      if(mode!=='DIAG' && pending.length > 0) {
        pending.forEach(w => { tS.getRangeByIndexes(w.r, w.c, 1, 1).values = [[w.v]]; });
        summary.innerHTML = `✅ ${pending.length}건 성공! ` + (results.length > 0 ? `<br><span style='color:red;'>(${results.length}건 누락)</span>` : "");
      } else { summary.innerHTML = results.length > 0 ? `🧐 미매칭 ${results.length}건 발견` : `✨ 완벽 매칭!`; }
      list.innerHTML = results.map(r => `<li onclick="window.goToCell('${cfg.srcS}', ${r.absRow})"><span class="acc-name">${r.name}</span><span class="diag-msg">❌ 짝 없음</span></li>`).join("");
      await context.sync();
    });
  } catch (e) { summary.innerText = "❌ 오류: " + e.message; }
}

async function runStep1Verify() {
  const summary = document.getElementById('summary-text');
  const tgtName = document.getElementById('s1-tgt-sheet').value;
  const tgtR = document.getElementById('s1-tgt-range').value;
  const cfg = { cur: parseInt(document.getElementById('s1-tgt-cur').value)-1, pri: parseInt(document.getElementById('s1-tgt-pri').value)-1, fin: parseInt(document.getElementById('s1-tgt-fin').value)-1 };
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(tgtName);
      const range = sheet.getRange(tgtR).load(["formulas", "columnIndex", "rowIndex"]);
      await context.sync();
      const tSC = range.columnIndex; const tSR = range.rowIndex;
      const finL = getColLetter(tSC + cfg.fin); const curL = getColLetter(tSC + cfg.cur); const priL = getColLetter(tSC + cfg.pri);
      let count = 0; const regex = new RegExp(`${finL}(\\d+)`, 'g');
      range.formulas.forEach((row, rIdx) => {
        const f = row[cfg.fin];
        if(f && typeof f === "string" && f.startsWith("=") && f.includes(finL)) {
          sheet.getCell(tSR + rIdx, tSC + cfg.cur).formulas = [[f.replace(regex, `${curL}$1`)]];
          sheet.getCell(tSR + rIdx, tSC + cfg.pri).formulas = [[f.replace(regex, `${priL}$1`)]];
          count++;
        }
      });
      summary.innerHTML = `✅ ${count}개 수식 복구 완료`; await context.sync();
    });
  } catch (e) { summary.innerText = "❌ 오류"; }
}

/* ──────── [2단계] 공시 전송 (🚀 셀 참조 수식 전송 보완) ──────── */

async function runStep2(mode) {
  const summary = document.getElementById('summary-text'); const list = document.getElementById('mismatch-list');
  const cfg = {
    srcS: document.getElementById('s2-src-sheet').value, srcR: document.getElementById('s2-src-range').value,
    tgtS: document.getElementById('s2-tgt-sheet').value, tgtR: document.getElementById('s2-tgt-range').value,
    srcAcc: parseInt(document.getElementById('s2-src-acc').value)-1, 
    srcPri: parseInt(document.getElementById('s2-src-pri').value)-1, 
    srcCur: parseInt(document.getElementById('s2-src-cur').value)-1,
    tgtAcc: parseInt(document.getElementById('s2-tgt-acc').value)-1
  };
  try {
    summary.innerText = "2단계 매칭 및 수식 연결 중..."; list.innerHTML = "";
    await Excel.run(async (context) => {
      const sS = context.workbook.worksheets.getItem(cfg.srcS); const tS = context.workbook.worksheets.getItem(cfg.tgtS);
      const sRO = sS.getRange(cfg.srcR).load(["values", "rowIndex", "columnIndex"]); // 🚀 columnIndex 추가 로드
      const tRO = tS.getRange(cfg.tgtR).load(["values", "rowIndex", "columnIndex"]);
      await context.sync();
      const tSC = tRO.columnIndex; const tSR = tRO.rowIndex;

      const srcMap = new Map(); let lastSrcMain = "Root";
      const colLetCur = getColLetter(sRO.columnIndex + cfg.srcCur);
      const colLetPri = getColLetter(sRO.columnIndex + cfg.srcPri);

      sRO.values.forEach((r, idx) => {
        const raw = r[cfg.srcAcc]; if(!raw) return;
        const cn = superClean(raw); const isH = isSumOrHeader(raw);
        const effectiveParent = !isH && contraKeywords.some(k => cn.includes(k)) ? lastSrcMain : "Root";
        const key = `${effectiveParent}_${cn}`;
        
        // 🚀 수식 생성 (절대 주소 형태)
        const excelRow = sRO.rowIndex + idx + 1;
        const curFormula = `='${cfg.srcS}'!$${colLetCur}$${excelRow}`;
        const priFormula = `='${cfg.srcS}'!$${colLetPri}$${excelRow}`;

        if(!srcMap.has(key)) srcMap.set(key, []);
        srcMap.get(key).push({ 
            cVal: parseNum(r[cfg.srcCur]), pVal: parseNum(r[cfg.srcPri]), // 진단용 값
            cRef: curFormula, pRef: priFormula,                          // 🚀 전송용 수식
            idx: idx, cn: cn, raw: raw 
        });
        if(!isH) lastSrcMain = cn;
      });

      const pending = [], results = [], srcUsed = new Set(), tgtUseCount = {};
      let lastTgtMain = "Root";

      tRO.values.forEach((row, i) => {
        const raw = row[cfg.tgtAcc]; if(!raw) return;
        const cn = superClean(raw); const isH = isSumOrHeader(raw);
        const effectiveParent = !isH && contraKeywords.some(k => cn.includes(k)) ? lastTgtMain : "Root";
        const key = `${effectiveParent}_${cn}`;

        let m = null;
        if(srcMap.has(key)) {
            const matches = srcMap.get(key); const use = tgtUseCount[key] || 0;
            if(matches[use]) { m = matches[use]; tgtUseCount[key] = use + 1; }
        }
        if(!m && effectiveParent !== "Root") {
            srcMap.forEach((val, k) => { if(k.startsWith(effectiveParent + "_")) {
                const found = val.find(item => item.cn.includes(cn) || cn.includes(item.cn));
                if(found) m = found;
            }});
        }

        if(m) {
          srcUsed.add(m.idx);
          if(mode !== 'DIAG') {
            const isH_tgt = isSumOrHeader(raw);
            if(mode==='SYNC_CUR') pending.push({r: tSR + i, c: tSC + cfg.tgtAcc + (isH_tgt ? 2 : 1), f: m.cRef});
            if(mode==='SYNC_PRI') pending.push({r: tSR + i, c: tSC + cfg.tgtAcc + (isH_tgt ? 4 : 3), f: m.pRef});
          }
        }
        if(!isH) lastTgtMain = cn;
      });

      if(mode === 'DIAG') {
          sRO.values.forEach((r, idx) => {
              const raw = r[cfg.srcAcc]; if(!raw || isSumOrHeader(raw)) return;
              if((parseNum(r[cfg.srcCur]) !== 0 || parseNum(r[cfg.srcPri]) !== 0) && !srcUsed.has(idx)) {
                  results.push({ name: raw, absRow: sRO.rowIndex + idx, msg: "❌ 공시용 누락 (금액 존재)" });
              }
          });
      }

      if(mode !== 'DIAG' && pending.length > 0) {
        // 🚀 수식으로 전송 (.formulas 속성 사용)
        pending.forEach(w => { tS.getRangeByIndexes(w.r, w.c, 1, 1).formulas = [[w.f]]; });
        summary.innerHTML = `✅ ${pending.length}건 수식 연결 성공!`;
      } else {
        summary.innerHTML = results.length > 0 ? `🧐 미매칭 ${results.length}건 발견` : `✨ 연동 완료!`;
      }
      
      list.innerHTML = results.map(r => `<li onclick="window.goToCell('${cfg.srcS}', ${r.absRow})"><span class="acc-name">${r.name}</span><span class="diag-msg">${r.msg}</span></li>`).join("");
      await context.sync();
    });
  } catch (e) { summary.innerText = "❌ 오류: " + e.message; }
}

async function runStep2Verify() {
  const summary = document.getElementById('summary-text');
  const tgtS = document.getElementById('s2-tgt-sheet').value;
  const tgtR = document.getElementById('s2-tgt-range').value;
  const accCol = parseInt(document.getElementById('s2-tgt-acc').value)-1;
  try {
    summary.innerText = "계층 수식 보정 중...";
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(tgtS);
      const range = sheet.getRange(tgtR).load(["values", "rowIndex", "columnIndex"]);
      await context.sync();
      const tSC = range.columnIndex; const tSR = range.rowIndex;
      const colCL = getColLetter(tSC + accCol + 1); const colDL = getColLetter(tSC + accCol + 2);
      const colEL = getColLetter(tSC + accCol + 3); const colFL = getColLetter(tSC + accCol + 4);
      let romans = [], lastRoman = -1, lastSub = -1, count = 0;
      const pS = (s, e) => {
        if (s === -1 || s >= e) return;
        const start = tSR + s + 2, end = tSR + e;
        sheet.getCell(tSR + s, tSC + accCol + 2).formulas = [[`=SUM(${colCL}${start}:${colCL}${end})`]];
        sheet.getCell(tSR + s, tSC + accCol + 4).formulas = [[`=SUM(${colEL}${start}:${colEL}${end})`]];
        count += 2;
      };
      const pR = (r, n) => {
        if (r === -1) return;
        let subC = [], subP = [];
        for (let k = r + 1; k < n; k++) { if (isSubHeader(range.values[k][accCol])) { subC.push(`${colDL}${tSR + k + 1}`); subP.push(`${colFL}${tSR + k + 1}`); } }
        if (subC.length > 0) {
          sheet.getCell(tSR + r, tSC + accCol + 2).formulas = [[`=${subC.join("+")}`]];
          sheet.getCell(tSR + r, tSC + accCol + 4).formulas = [[`=${subP.join("+")}`]];
        } else { pS(r, n); }
        count += 2;
      };
      range.values.forEach((row, idx) => {
        const name = String(row[accCol] || "").trim();
        if (isRoman(name)) { if (lastSub !== -1) pS(lastSub, idx); if (lastRoman !== -1) pR(lastRoman, idx); romans.push(idx); lastRoman = idx; lastSub = -1; }
        else if (isSubHeader(name)) { if (lastSub !== -1) pS(lastSub, idx); lastSub = idx; }
        else if (isGrandTotal(name)) {
          if (lastSub !== -1) pS(lastSub, idx); if (lastRoman !== -1) pR(lastRoman, idx);
          if (romans.length > 0) {
            sheet.getCell(tSR + idx, tSC + accCol + 2).formulas = [[`=${romans.map(i => `${colDL}${tSR + i + 1}`).join("+")}`]];
            sheet.getCell(tSR + idx, tSC + accCol + 4).formulas = [[`=${romans.map(i => `${colFL}${tSR + i + 1}`).join("+")}`]];
          }
          romans = []; lastRoman = -1; lastSub = -1; count += 2;
        }
      });
      summary.innerHTML = `✅ ${count}개 수식 복구 성공`; await context.sync();
    });
  } catch (e) { summary.innerText = "❌ 실패: " + e.message; }
}