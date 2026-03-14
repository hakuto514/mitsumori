/* ===== V3 Excel Export with Styled Borders (xlsx-js-style) ===== */
/* V3変更点:
 *  - 日付フォーマット: toJpDateStr() で YYYY年MM月DD日 に統一（PDFと一致）
 *  - 金額セル: setCellMoney() で numFmt "#,##0" を適用（カンマ区切り統一）
 */

function exportExcel() {
  const wb = XLSX.utils.book_new();
  const aux = Number(document.getElementById("auxMaterial").value) || 0;
  let sellTotal = 0,
    costTotal = 0;
  parts.forEach((p) => {
    const s = floor100(p.sellInput);
    sellTotal += s * p.qty;
    costTotal +=
      (p.source === "self" ? Math.round(s * 0.62) : p.costPrice) * p.qty;
  });
  let laborSum = 0;
  labors.forEach((l) => (laborSum += l.amount));
  let discSum = 0;
  discounts.forEach((d) => (discSum += d.amount));
  const subtotal = sellTotal + aux + laborSum - discSum;
  const tax = Math.floor(subtotal * 0.1);
  const total = subtotal + tax;
  const staff = staffList[selectedStaff] || "";
  const dateVal = document.getElementById("issueDate").value;
  const installDateVal = document.getElementById("installDate").value;

  // V3: toJpDateStr() で日付フォーマットを YYYY年MM月DD日 に統一
  const dFmt = toJpDateStr(dateVal);
  const iDFmt = toJpDateStr(installDateVal);

  // V3: 金額フォーマット定数
  const FMT_JPY = "#,##0";

  // Style Definitions
  const thinBorder = { style: "thin", color: { rgb: "000000" } };
  const allBorders = {
    top: thinBorder,
    bottom: thinBorder,
    left: thinBorder,
    right: thinBorder,
  };
  const fontNormal = { sz: 10, name: "ＭＳ Ｐゴシック" };
  const fontSmall = { sz: 9, name: "ＭＳ Ｐゴシック" };
  const fontLarge = { sz: 16, name: "ＭＳ Ｐゴシック", bold: true };

  const alignLC = { vertical: "center", horizontal: "left", wrapText: true };
  const alignCC = { vertical: "center", horizontal: "center", wrapText: true };
  const alignRC = { vertical: "center", horizontal: "right", wrapText: true };

  const styleNormal = { font: fontNormal, alignment: alignLC };
  const styleCenter = { font: fontNormal, alignment: alignCC };
  const styleRight = { font: fontNormal, alignment: alignRC };

  const MAX_ROWS = 64;
  const MAX_COLS = 32;
  const rows = [];
  const styles = [];
  for (let r = 0; r < MAX_ROWS; r++) {
    const row = [];
    const rowStyles = [];
    for (let c = 0; c < MAX_COLS; c++) {
      row.push("");
      rowStyles.push({ ...styleNormal });
    }
    rows.push(row);
    styles.push(rowStyles);
  }

  function setCell(
    r,
    c,
    val,
    align = "left",
    customFont = fontNormal,
    border = null,
    numFmt = null,
  ) {
    if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) return;
    rows[r][c] = val;
    const s = { font: customFont };
    if (align === "center") s.alignment = alignCC;
    else if (align === "right") s.alignment = alignRC;
    else s.alignment = alignLC;
    if (border) s.border = border;
    // V3: numFmt を style に追加（xlsx-js-style対応）
    if (numFmt) s.numFmt = numFmt;
    styles[r][c] = s;
  }

  // V3: 金額セル専用ヘルパー（numFmt "#,##0" 適用）
  function setCellMoney(r, c, val, border = null) {
    setCell(r, c, val, "right", fontNormal, border, FMT_JPY);
  }

  function setRangeBorder(r1, c1, r2, c2, border = allBorders) {
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        let cellBorder = {};
        if (r === r1) cellBorder.top = border.top;
        if (r === r2) cellBorder.bottom = border.bottom;
        if (c === c1) cellBorder.left = border.left;
        if (c === c2) cellBorder.right = border.right;
        if (!styles[r][c].border) styles[r][c].border = {};
        Object.assign(styles[r][c].border, cellBorder);
      }
    }
  }

  // 1. Header Information
  setCell(0, 12, "御　見　積　書", "center", fontLarge);
  setCell(0, 22, "見積書番号", "center", fontSmall);
  setCell(0, 26, document.getElementById("estNo").value, "center", fontSmall);

  setCell(3, 0, (document.getElementById("custName").value || "") + "　御中", "left", { sz: 14, name: "ＭＳ Ｐゴシック", bold: true });
  // V3: 日付フォーマット統一
  setCell(3, 23, dFmt, "right", fontNormal);

  setCell(6, 0, "下記の通り御見積申し上げます。");
  setCell(7, 0, "何卒ご用命下さいますようお願い申し上げます。");
  setCell(7, 20, document.getElementById("compName").value);
  setCell(8, 20, document.getElementById("compAddr").value, "left", fontSmall);

  setCell(8, 0, "納入場所", "center");
  setCell(8, 5, document.getElementById("delivPlace").value);
  setCell(10, 0, "納期", "center");
  setCell(10, 5, document.getElementById("delivDate").value);
  setCell(11, 0, "御支払条件", "center");
  setCell(11, 5, document.getElementById("payTerms").value);
  setCell(12, 0, "見積有効期限", "center");
  setCell(12, 5, document.getElementById("validity").value);

  setCell(8, 15, "様", "center");
  setCell(8, 19, "担当部署", "center");
  setCell(8, 25, document.getElementById("compDept").value);
  setCell(10, 19, "TEL", "center");
  setCell(10, 20, document.getElementById("compTel").value);
  setCell(10, 26, "FAX ", "center");
  setCell(10, 27, document.getElementById("compFax").value);
  setCell(11, 21, "担当者", "center");
  setCell(11, 24, staff);
  setCell(13, 26, "上長", "center");
  setCell(13, 29, "担当者", "center");

  // 2. Equipment Information
  setCell(14, 0, "製品名", "center");
  setCell(14, 4, document.getElementById("prodName").value);
  setCell(14, 12, "機種", "center");
  setCell(14, 15, document.getElementById("modelNo").value);
  setCell(15, 0, "設置日", "center");
  // V3: 設置日も統一フォーマット
  setCell(15, 4, iDFmt);
  setCell(15, 12, "機番", "center");
  setCell(15, 15, document.getElementById("serialNo").value);

  // 3. Call Content
  setCell(16, 0, "【コール内容】");
  setCell(16, 5, document.getElementById("callContent").value);
  setCell(17, 0, "【故障原因・処置】");
  setCell(17, 5, document.getElementById("causeAction").value);

  setCell(20, 0, "故障箇所の部品のみの修繕内容です。", "left", { sz: 9, name: "ＭＳ Ｐゴシック", bold: true });

  // 4. Parts Table
  setCell(21, 0, "No", "center");
  setRangeBorder(21, 0, 21, 0);
  setCell(21, 1, "部品名", "center");
  setRangeBorder(21, 1, 21, 12);
  setCell(21, 13, "部品コード", "center");
  setRangeBorder(21, 13, 21, 18);
  setCell(21, 19, "数量", "center");
  setRangeBorder(21, 19, 21, 21);
  setCell(21, 22, "単価", "center");
  setRangeBorder(21, 22, 21, 26);
  setCell(21, 27, "金額", "center");
  setRangeBorder(21, 27, 21, 31);

  let rIdx = 22;
  for (let i = 0; i < 15; i++) {
    setRangeBorder(rIdx + i, 0, rIdx + i, 0);
    setRangeBorder(rIdx + i, 1, rIdx + i, 12);
    setRangeBorder(rIdx + i, 13, rIdx + i, 18);
    setRangeBorder(rIdx + i, 19, rIdx + i, 21);
    setRangeBorder(rIdx + i, 22, rIdx + i, 26);
    setRangeBorder(rIdx + i, 27, rIdx + i, 31);
    setCell(rIdx + i, 0, i + 1, "center");
  }

  for (let i = 0; i < parts.length && i < 15; i++) {
    const pt = parts[i];
    const sell = floor100(pt.sellInput);
    setCell(rIdx + i, 1, pt.name, "left");
    setCell(rIdx + i, 13, pt.code, "left");
    setCell(rIdx + i, 19, pt.qty || "", "center");
    // V3: setCellMoney で金額セルにカンマ書式を適用
    if (pt.name) {
      setCellMoney(rIdx + i, 22, sell === 0 ? 0 : sell);
      setCellMoney(rIdx + i, 27, sell * pt.qty === 0 ? 0 : sell * pt.qty);
    }
  }

  // 5. Labor & Totals
  setCell(37, 0, "補修・補助材料費", "center");
  setRangeBorder(37, 0, 37, 19);
  setCell(37, 20, aux > 0 ? "一式" : "", "center");
  setRangeBorder(37, 20, 37, 21);
  // V3: setCellMoney
  setCellMoney(37, 27, aux || 0);
  setRangeBorder(37, 22, 37, 31);

  setCell(38, 0, "使用部品費合計", "center");
  setRangeBorder(38, 0, 38, 19);
  // V3: setCellMoney
  setCellMoney(38, 22, sellTotal + aux);
  setRangeBorder(38, 22, 38, 31);

  setCell(39, 0, "作業名　及び　値引き", "center");
  setRangeBorder(39, 0, 49, 1);

  let lIdx = 39;
  labors.forEach((l) => {
    if (lIdx <= 49) {
      setCell(lIdx, 2, l.name, "left");
      setRangeBorder(lIdx, 2, lIdx, 18);
      setCell(lIdx, 19, l.amount > 0 ? "一式" : "", "center");
      setRangeBorder(lIdx, 19, lIdx, 21);
      // V3: setCellMoney
      setCellMoney(lIdx, 22, l.amount || 0);
      setRangeBorder(lIdx, 22, lIdx, 31);
      lIdx++;
    }
  });
  discounts.forEach((d) => {
    if (lIdx <= 49) {
      setCell(lIdx, 2, d.name, "left");
      setRangeBorder(lIdx, 2, lIdx, 18);
      setCell(lIdx, 19, "", "center");
      setRangeBorder(lIdx, 19, lIdx, 21);
      // V3: setCellMoney（値引きはマイナス）
      setCellMoney(lIdx, 22, d.amount ? -d.amount : 0);
      setRangeBorder(lIdx, 22, lIdx, 31);
      lIdx++;
    }
  });
  for (; lIdx <= 49; lIdx++) {
    setRangeBorder(lIdx, 2, lIdx, 18);
    setRangeBorder(lIdx, 19, lIdx, 21);
    setRangeBorder(lIdx, 22, lIdx, 31);
  }

  setCell(50, 19, "小計", "center");
  setRangeBorder(50, 19, 50, 23);
  // V3: setCellMoney
  setCellMoney(50, 24, subtotal);
  setRangeBorder(50, 24, 50, 31);

  setCell(51, 19, "消費税", "center");
  setRangeBorder(51, 19, 51, 23);
  setCellMoney(51, 24, tax);
  setRangeBorder(51, 24, 51, 31);

  setCell(52, 19, "合計", "center");
  setRangeBorder(52, 19, 53, 23);
  setCellMoney(52, 24, total);
  setRangeBorder(52, 24, 53, 31);

  // 6. Reply section
  setCell(56, 0, "御返答欄", "center");
  setRangeBorder(56, 0, 61, 1);
  setCell(56, 2, "上記御見積書の御返答を御願い致します。", "left");

  setCell(58, 3, "１. 修理する");
  setCell(58, 16, "御返答日");
  setCell(58, 22, "年", "center");
  setCell(58, 25, "月", "center");
  setCell(58, 28, "日", "center");

  setCell(60, 3, "２. 修理しない");
  setCell(60, 16, "発注者");
  setCell(60, 25, "様");
  setCell(60, 27, "印", "center");

  setRangeBorder(56, 0, 61, 31, {
    top: thinBorder,
    bottom: thinBorder,
    left: thinBorder,
    right: thinBorder,
  });

  // 7. Remarks
  const remarksText = document.getElementById("remarks").value;
  if (remarksText) {
    const remarks = remarksText.split("\n");
    if (remarks[0]) {
      setCell(62, 0, "＊", "right", fontSmall);
      setCell(62, 1, remarks[0], "left", fontSmall);
    }
    if (remarks[1]) {
      setCell(63, 0, "＊", "right", fontSmall);
      setCell(63, 1, remarks[1], "left", fontSmall);
    }
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);

  // Apply Styles
  for (let r = 0; r < MAX_ROWS; r++) {
    for (let c = 0; c < MAX_COLS; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (!ws[addr]) ws[addr] = { v: "", t: "s" };
      ws[addr].s = styles[r][c];
      // V3: numFmt を ws[addr].z にも設定（SheetJS互換）
      if (styles[r][c].numFmt) {
        ws[addr].z = styles[r][c].numFmt;
      }
    }
  }

  ws["!merges"] = [{"s": {"r": 36, "c": 22}, "e": {"r": 36, "c": 26}}, {"s": {"r": 40, "c": 22}, "e": {"r": 40, "c": 31}}, {"s": {"r": 36, "c": 13}, "e": {"r": 36, "c": 18}}, {"s": {"r": 34, "c": 13}, "e": {"r": 34, "c": 18}}, {"s": {"r": 36, "c": 1}, "e": {"r": 36, "c": 12}}, {"s": {"r": 63, "c": 1}, "e": {"r": 63, "c": 33}}, {"s": {"r": 38, "c": 0}, "e": {"r": 38, "c": 19}}, {"s": {"r": 39, "c": 0}, "e": {"r": 49, "c": 1}}, {"s": {"r": 45, "c": 22}, "e": {"r": 45, "c": 31}}, {"s": {"r": 44, "c": 22}, "e": {"r": 44, "c": 31}}, {"s": {"r": 38, "c": 22}, "e": {"r": 38, "c": 31}}, {"s": {"r": 15, "c": 4}, "e": {"r": 15, "c": 11}}, {"s": {"r": 8, "c": 0}, "e": {"r": 8, "c": 4}}, {"s": {"r": 13, "c": 29}, "e": {"r": 13, "c": 31}}, {"s": {"r": 14, "c": 0}, "e": {"r": 14, "c": 3}}, {"s": {"r": 10, "c": 0}, "e": {"r": 10, "c": 4}}, {"s": {"r": 10, "c": 5}, "e": {"r": 10, "c": 9}}, {"s": {"r": 10, "c": 15}, "e": {"r": 10, "c": 16}}, {"s": {"r": 3, "c": 0}, "e": {"r": 4, "c": 19}}, {"s": {"r": 8, "c": 25}, "e": {"r": 8, "c": 31}}, {"s": {"r": 7, "c": 20}, "e": {"r": 7, "c": 31}}, {"s": {"r": 56, "c": 0}, "e": {"r": 61, "c": 0}}, {"s": {"r": 10, "c": 27}, "e": {"r": 10, "c": 31}}, {"s": {"r": 11, "c": 0}, "e": {"r": 11, "c": 4}}, {"s": {"r": 11, "c": 5}, "e": {"r": 11, "c": 9}}, {"s": {"r": 37, "c": 27}, "e": {"r": 37, "c": 31}}, {"s": {"r": 46, "c": 22}, "e": {"r": 46, "c": 31}}, {"s": {"r": 34, "c": 27}, "e": {"r": 34, "c": 31}}, {"s": {"r": 33, "c": 13}, "e": {"r": 33, "c": 18}}, {"s": {"r": 12, "c": 0}, "e": {"r": 12, "c": 4}}, {"s": {"r": 14, "c": 4}, "e": {"r": 14, "c": 11}}, {"s": {"r": 14, "c": 12}, "e": {"r": 14, "c": 14}}, {"s": {"r": 47, "c": 22}, "e": {"r": 47, "c": 31}}, {"s": {"r": 48, "c": 22}, "e": {"r": 48, "c": 31}}, {"s": {"r": 49, "c": 22}, "e": {"r": 49, "c": 31}}, {"s": {"r": 43, "c": 22}, "e": {"r": 43, "c": 31}}, {"s": {"r": 35, "c": 19}, "e": {"r": 35, "c": 20}}, {"s": {"r": 35, "c": 22}, "e": {"r": 35, "c": 26}}, {"s": {"r": 37, "c": 0}, "e": {"r": 37, "c": 19}}, {"s": {"r": 31, "c": 22}, "e": {"r": 31, "c": 26}}, {"s": {"r": 32, "c": 22}, "e": {"r": 32, "c": 26}}, {"s": {"r": 31, "c": 19}, "e": {"r": 31, "c": 20}}, {"s": {"r": 32, "c": 19}, "e": {"r": 32, "c": 20}}, {"s": {"r": 33, "c": 19}, "e": {"r": 33, "c": 20}}, {"s": {"r": 33, "c": 22}, "e": {"r": 33, "c": 26}}, {"s": {"r": 27, "c": 22}, "e": {"r": 27, "c": 26}}, {"s": {"r": 24, "c": 27}, "e": {"r": 24, "c": 31}}, {"s": {"r": 25, "c": 27}, "e": {"r": 25, "c": 31}}, {"s": {"r": 26, "c": 27}, "e": {"r": 26, "c": 31}}, {"s": {"r": 27, "c": 27}, "e": {"r": 27, "c": 31}}, {"s": {"r": 28, "c": 22}, "e": {"r": 28, "c": 26}}, {"s": {"r": 28, "c": 27}, "e": {"r": 28, "c": 31}}, {"s": {"r": 4, "c": 23}, "e": {"r": 4, "c": 31}}, {"s": {"r": 19, "c": 5}, "e": {"r": 19, "c": 28}}, {"s": {"r": 28, "c": 13}, "e": {"r": 28, "c": 18}}, {"s": {"r": 29, "c": 13}, "e": {"r": 29, "c": 18}}, {"s": {"r": 25, "c": 1}, "e": {"r": 25, "c": 12}}, {"s": {"r": 26, "c": 1}, "e": {"r": 26, "c": 12}}, {"s": {"r": 27, "c": 1}, "e": {"r": 27, "c": 12}}, {"s": {"r": 28, "c": 1}, "e": {"r": 28, "c": 12}}, {"s": {"r": 29, "c": 1}, "e": {"r": 29, "c": 12}}, {"s": {"r": 15, "c": 0}, "e": {"r": 15, "c": 3}}, {"s": {"r": 30, "c": 1}, "e": {"r": 30, "c": 12}}, {"s": {"r": 31, "c": 1}, "e": {"r": 31, "c": 12}}, {"s": {"r": 32, "c": 1}, "e": {"r": 32, "c": 12}}, {"s": {"r": 33, "c": 1}, "e": {"r": 33, "c": 12}}, {"s": {"r": 17, "c": 5}, "e": {"r": 17, "c": 28}}, {"s": {"r": 23, "c": 22}, "e": {"r": 23, "c": 26}}, {"s": {"r": 24, "c": 22}, "e": {"r": 24, "c": 26}}, {"s": {"r": 25, "c": 22}, "e": {"r": 25, "c": 26}}, {"s": {"r": 26, "c": 22}, "e": {"r": 26, "c": 26}}, {"s": {"r": 25, "c": 13}, "e": {"r": 25, "c": 18}}, {"s": {"r": 17, "c": 0}, "e": {"r": 19, "c": 4}}, {"s": {"r": 22, "c": 1}, "e": {"r": 22, "c": 12}}, {"s": {"r": 23, "c": 1}, "e": {"r": 23, "c": 12}}, {"s": {"r": 24, "c": 1}, "e": {"r": 24, "c": 12}}, {"s": {"r": 24, "c": 13}, "e": {"r": 24, "c": 18}}, {"s": {"r": 21, "c": 19}, "e": {"r": 21, "c": 21}}, {"s": {"r": 8, "c": 5}, "e": {"r": 8, "c": 14}}, {"s": {"r": 8, "c": 15}, "e": {"r": 8, "c": 16}}, {"s": {"r": 9, "c": 18}, "e": {"r": 9, "c": 22}}, {"s": {"r": 11, "c": 21}, "e": {"r": 11, "c": 23}}, {"s": {"r": 12, "c": 5}, "e": {"r": 12, "c": 9}}, {"s": {"r": 10, "c": 20}, "e": {"r": 10, "c": 24}}, {"s": {"r": 9, "c": 23}, "e": {"r": 9, "c": 30}}, {"s": {"r": 8, "c": 19}, "e": {"r": 8, "c": 24}}, {"s": {"r": 26, "c": 13}, "e": {"r": 26, "c": 18}}, {"s": {"r": 20, "c": 0}, "e": {"r": 20, "c": 31}}, {"s": {"r": 11, "c": 24}, "e": {"r": 11, "c": 31}}, {"s": {"r": 21, "c": 1}, "e": {"r": 21, "c": 12}}, {"s": {"r": 21, "c": 22}, "e": {"r": 21, "c": 26}}, {"s": {"r": 21, "c": 27}, "e": {"r": 21, "c": 31}}, {"s": {"r": 16, "c": 0}, "e": {"r": 16, "c": 4}}, {"s": {"r": 22, "c": 27}, "e": {"r": 22, "c": 31}}, {"s": {"r": 23, "c": 27}, "e": {"r": 23, "c": 31}}, {"s": {"r": 22, "c": 22}, "e": {"r": 22, "c": 26}}, {"s": {"r": 13, "c": 23}, "e": {"r": 13, "c": 25}}, {"s": {"r": 23, "c": 13}, "e": {"r": 23, "c": 18}}, {"s": {"r": 16, "c": 5}, "e": {"r": 16, "c": 21}}, {"s": {"r": 15, "c": 12}, "e": {"r": 15, "c": 14}}, {"s": {"r": 18, "c": 5}, "e": {"r": 18, "c": 28}}, {"s": {"r": 13, "c": 26}, "e": {"r": 13, "c": 28}}, {"s": {"r": 14, "c": 15}, "e": {"r": 14, "c": 21}}, {"s": {"r": 15, "c": 15}, "e": {"r": 15, "c": 21}}, {"s": {"r": 21, "c": 13}, "e": {"r": 21, "c": 18}}, {"s": {"r": 22, "c": 13}, "e": {"r": 22, "c": 18}}, {"s": {"r": 36, "c": 27}, "e": {"r": 36, "c": 31}}, {"s": {"r": 29, "c": 22}, "e": {"r": 29, "c": 26}}, {"s": {"r": 30, "c": 22}, "e": {"r": 30, "c": 26}}, {"s": {"r": 29, "c": 27}, "e": {"r": 29, "c": 31}}, {"s": {"r": 33, "c": 27}, "e": {"r": 33, "c": 31}}, {"s": {"r": 37, "c": 20}, "e": {"r": 37, "c": 21}}, {"s": {"r": 41, "c": 22}, "e": {"r": 41, "c": 31}}, {"s": {"r": 31, "c": 27}, "e": {"r": 31, "c": 31}}, {"s": {"r": 32, "c": 27}, "e": {"r": 32, "c": 31}}, {"s": {"r": 30, "c": 27}, "e": {"r": 30, "c": 31}}, {"s": {"r": 34, "c": 22}, "e": {"r": 34, "c": 26}}, {"s": {"r": 27, "c": 19}, "e": {"r": 27, "c": 20}}, {"s": {"r": 28, "c": 19}, "e": {"r": 28, "c": 20}}, {"s": {"r": 29, "c": 19}, "e": {"r": 29, "c": 20}}, {"s": {"r": 30, "c": 19}, "e": {"r": 30, "c": 20}}, {"s": {"r": 34, "c": 1}, "e": {"r": 34, "c": 12}}, {"s": {"r": 35, "c": 1}, "e": {"r": 35, "c": 12}}, {"s": {"r": 27, "c": 13}, "e": {"r": 27, "c": 18}}, {"s": {"r": 31, "c": 13}, "e": {"r": 31, "c": 18}}, {"s": {"r": 32, "c": 13}, "e": {"r": 32, "c": 18}}, {"s": {"r": 35, "c": 13}, "e": {"r": 35, "c": 18}}, {"s": {"r": 30, "c": 13}, "e": {"r": 30, "c": 18}}, {"s": {"r": 34, "c": 19}, "e": {"r": 34, "c": 20}}, {"s": {"r": 0, "c": 22}, "e": {"r": 0, "c": 25}}, {"s": {"r": 0, "c": 26}, "e": {"r": 0, "c": 31}}, {"s": {"r": 0, "c": 12}, "e": {"r": 1, "c": 19}}, {"s": {"r": 36, "c": 19}, "e": {"r": 36, "c": 20}}, {"s": {"r": 22, "c": 19}, "e": {"r": 22, "c": 20}}, {"s": {"r": 23, "c": 19}, "e": {"r": 23, "c": 20}}, {"s": {"r": 24, "c": 19}, "e": {"r": 24, "c": 20}}, {"s": {"r": 25, "c": 19}, "e": {"r": 25, "c": 20}}, {"s": {"r": 26, "c": 19}, "e": {"r": 26, "c": 20}}, {"s": {"r": 39, "c": 2}, "e": {"r": 39, "c": 18}}, {"s": {"r": 40, "c": 2}, "e": {"r": 40, "c": 18}}, {"s": {"r": 41, "c": 2}, "e": {"r": 41, "c": 18}}, {"s": {"r": 42, "c": 2}, "e": {"r": 42, "c": 18}}, {"s": {"r": 43, "c": 2}, "e": {"r": 43, "c": 18}}, {"s": {"r": 44, "c": 2}, "e": {"r": 44, "c": 18}}, {"s": {"r": 45, "c": 2}, "e": {"r": 45, "c": 18}}, {"s": {"r": 46, "c": 2}, "e": {"r": 46, "c": 18}}, {"s": {"r": 47, "c": 2}, "e": {"r": 47, "c": 18}}, {"s": {"r": 48, "c": 2}, "e": {"r": 48, "c": 18}}, {"s": {"r": 49, "c": 2}, "e": {"r": 49, "c": 18}}, {"s": {"r": 39, "c": 19}, "e": {"r": 39, "c": 21}}, {"s": {"r": 40, "c": 19}, "e": {"r": 40, "c": 21}}, {"s": {"r": 41, "c": 19}, "e": {"r": 41, "c": 21}}, {"s": {"r": 42, "c": 19}, "e": {"r": 42, "c": 21}}, {"s": {"r": 43, "c": 19}, "e": {"r": 43, "c": 21}}, {"s": {"r": 44, "c": 19}, "e": {"r": 44, "c": 21}}, {"s": {"r": 45, "c": 19}, "e": {"r": 45, "c": 21}}, {"s": {"r": 46, "c": 19}, "e": {"r": 46, "c": 21}}, {"s": {"r": 47, "c": 19}, "e": {"r": 47, "c": 21}}, {"s": {"r": 48, "c": 19}, "e": {"r": 48, "c": 21}}, {"s": {"r": 49, "c": 19}, "e": {"r": 49, "c": 21}}, {"s": {"r": 52, "c": 19}, "e": {"r": 53, "c": 23}}, {"s": {"r": 51, "c": 19}, "e": {"r": 51, "c": 23}}, {"s": {"r": 50, "c": 19}, "e": {"r": 50, "c": 23}}, {"s": {"r": 50, "c": 24}, "e": {"r": 50, "c": 31}}, {"s": {"r": 51, "c": 24}, "e": {"r": 51, "c": 31}}, {"s": {"r": 52, "c": 24}, "e": {"r": 53, "c": 31}}];

  const cols = [];
  for (let i = 0; i < MAX_COLS; i++) cols.push({ wch: 2.3 });
  ws["!cols"] = cols;

  const rowsHeight = [];
  for (let i = 0; i < MAX_ROWS; i++) rowsHeight.push({ hpt: 13.5 });
  rowsHeight[0] = { hpt: 20 };
  ws["!rows"] = rowsHeight;

  ws["!margins"] = { left: 0.3, right: 0.3, top: 0.3, bottom: 0.3, header: 0.15, footer: 0.15 };
  ws["!pageSetup"] = { paperSize: 9, orientation: "portrait", fitToWidth: 1, fitToHeight: 1 };

  if (!wb.Workbook) wb.Workbook = {};
  if (!wb.Workbook.Sheets) wb.Workbook.Sheets = [];
  wb.Workbook.Sheets[0] = { pageSetUpPr: { fitToPage: true } };

  XLSX.utils.book_append_sheet(wb, ws, "御見積書");
  XLSX.writeFile(wb, "御見積書_" + document.getElementById("estNo").value + ".xlsx");
  showToast("Excelを出力しました");
}
