/* 룰루피치 광고 판정 대시보드 — 공용 계산/렌더링 로직
   index.html(팀 전체 보기)과 admin.html(업로드) 양쪽에서 그대로 사용합니다. */

var DEFAULT_SETTINGS = {
  weightRoas: 0.7,
  weightInflowValue: 0.3,
  thFocus: 2.0,
  thKeep: 1.0,
  thCut: 0.7,
  cpaCapMultiplier: 3,
  inflowCostOutlierMultiplier: 3,
  noRevenueSpendMultiplier: 1.5,
  watchInflowCostRatio: 0.5,
  rankCap: 30,
  focusIncreasePct: 25,
  cutReducePct: 45,
  perCreativeCapPct: 50,
};

const BADGE_STYLES = {
  focus: { bg: "#dcfce7", fg: "#15803d", label: "🔼 집중" },
  keep: { bg: "#dbeafe", fg: "#1d4ed8", label: "➡️ 유지" },
  watch: { bg: "#fef9c3", fg: "#a16207", label: "🟡 소액관찰" },
  cut: { bg: "#ffedd5", fg: "#c2410c", label: "🔽 축소" },
  stop: { bg: "#fee2e2", fg: "#b91c1c", label: "⛔ 중단" },
  none: { bg: "#e2e8f0", fg: "#475569", label: "— 미분류" },
};

function fmtWon(n) {
  if (n === null || n === undefined || Number.isNaN(n)) return "-";
  return "₩" + Math.round(n).toLocaleString("ko-KR");
}
function fmtNum(n, digits) {
  digits = digits || 0;
  if (n === null || n === undefined || Number.isNaN(n)) return "-";
  return n.toLocaleString("ko-KR", { maximumFractionDigits: digits, minimumFractionDigits: digits });
}
function fmtPct(n, digits) {
  digits = digits || 0;
  if (n === null || n === undefined || Number.isNaN(n)) return "-";
  return (n >= 0 ? "+" : "") + n.toFixed(digits) + "%";
}
function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function extractCode(adName) {
  if (!adName) return null;
  const s = String(adName).trim();
  const idx = s.indexOf("_");
  return idx === -1 ? s : s.slice(0, idx);
}

function excelDateToStr(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return d.toISOString().slice(0, 10);
  }
  if (typeof v === "string") {
    const m = v.match(/\d{4}-\d{2}-\d{2}/);
    if (m) return m[0];
  }
  return null;
}

async function parseMetaFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
  return rows.map((r) => ({
    date: excelDateToStr(r["보고 시작"]),
    adName: r["광고 이름"],
    code: extractCode(r["광고 이름"]),
    spend: Number(r["지출 금액 (KRW)"]) || 0,
    purchases: Number(r["구매"]) || 0,
    purchaseValue: Number(r["구매 전환값"]) || 0,
  })).filter((r) => r.code);
}

function parseUtmFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      dynamicTyping: true,
      skipEmptyLines: true,
      complete: (res) => {
        resolve(res.data.map((r) => ({
          date: r.date,
          channel: r.channel ? String(r.channel).trim() : null,
          orderAmount: Number(r.order_amount) || 0,
          orderCount: Number(r.order_count) || 0,
          inflowCount: Number(r.inflow_count) || 0,
        })));
      },
      error: reject,
    });
  });
}

function dateRangeOf(rows) {
  const dates = rows.map((r) => r.date).filter(Boolean).sort();
  return dates.length ? [dates[0], dates[dates.length - 1]] : [null, null];
}
function inRange(date, start, end) {
  if (!date) return false;
  if (start && date < start) return false;
  if (end && date > end) return false;
  return true;
}
function classifyMedia(code) {
  if (!code) return "기타";
  if (/^SA_/i.test(code)) return "네이버검색";
  if (/^PMAX/i.test(code)) return "구글";
  if (/^(KKO|KAKAO)/i.test(code)) return "카카오";
  if (/^GD/i.test(code)) return "GFA/디스플레이";
  return "Meta";
}
function avg(arr) {
  const clean = arr.filter((n) => n !== null && n !== undefined && !Number.isNaN(n));
  if (!clean.length) return 0;
  return clean.reduce((s, n) => s + n, 0) / clean.length;
}

function computeDashboard(metaRows, utmRows, dateStart, dateEnd, settings) {
  const metaInRange = metaRows.filter((r) => inRange(r.date, dateStart, dateEnd));
  const utmInRange = utmRows.filter((r) => inRange(r.date, dateStart, dateEnd));

  const metaAgg = new Map();
  for (const r of metaInRange) {
    const cur = metaAgg.get(r.code) || { spend: 0, purchases: 0, purchaseValue: 0, adName: r.adName };
    cur.spend += r.spend;
    cur.purchases += r.purchases;
    cur.purchaseValue += r.purchaseValue;
    metaAgg.set(r.code, cur);
  }

  const utmAgg = new Map();
  let siteTotalRevenue = 0, siteTotalInflow = 0;
  const codeLike = /^\d{6,}/;
  for (const r of utmInRange) {
    siteTotalRevenue += r.orderAmount;
    siteTotalInflow += r.inflowCount;
    if (r.channel && codeLike.test(r.channel)) {
      const cur = utmAgg.get(r.channel) || { revenue: 0, orders: 0, inflow: 0 };
      cur.revenue += r.orderAmount;
      cur.orders += r.orderCount;
      cur.inflow += r.inflowCount;
      utmAgg.set(r.channel, cur);
    }
  }
  const siteAvgRevenuePerInflow = siteTotalInflow > 0 ? siteTotalRevenue / siteTotalInflow : 0;

  const allCodes = new Set([...metaAgg.keys(), ...utmAgg.keys()]);
  const creatives = [];
  const missingSpend = [];

  for (const code of allCodes) {
    const m = metaAgg.get(code);
    const u = utmAgg.get(code);
    if (!m) {
      missingSpend.push({ code, media: classifyMedia(code), inflow: u.inflow, orders: u.orders, revenue: u.revenue });
      continue;
    }
    const spend = m.spend;
    const inflow = u ? u.inflow : 0;
    const revenue = u ? u.revenue : 0;
    const orders = u ? u.orders : 0;
    const mediaRoas = spend > 0 && m.purchaseValue > 0 ? m.purchaseValue / spend : null;
    const inflowCost = inflow > 0 ? spend / inflow : null;
    const revenuePerInflow = inflow > 0 ? revenue / inflow : 0;
    const realRoas = spend > 0 ? revenue / spend : 0;
    const inflowValue = spend > 0 ? (inflow * siteAvgRevenuePerInflow) / spend : 0;
    const composite = realRoas * settings.weightRoas + inflowValue * settings.weightInflowValue;
    const cpa = orders > 0 ? spend / orders : null;
    creatives.push({ code, adName: m.adName, media: classifyMedia(code), spend, mediaRoas, inflow, inflowCost, revenuePerInflow, revenue, realRoas, inflowValue, composite, orders, cpa });
  }

  const spendCreatives = creatives.filter((c) => c.spend > 0);
  const avgInflowCost = avg(spendCreatives.filter((c) => c.inflowCost !== null).map((c) => c.inflowCost));
  const avgSpend = avg(spendCreatives.map((c) => c.spend));
  const avgCpa = avg(spendCreatives.filter((c) => c.cpa !== null).map((c) => c.cpa));

  const warnings = [];
  for (const c of creatives) {
    const reasons = [];
    let verdict;
    if (c.inflow === 0) {
      verdict = "stop"; reasons.push("유입 0 — UTM 세팅 점검");
    } else if (c.revenue === 0 && c.spend > avgSpend * settings.noRevenueSpendMultiplier) {
      verdict = "stop"; reasons.push("무매출 + 지출 과다");
    } else if (c.cpa !== null && avgCpa > 0 && c.cpa > avgCpa * settings.cpaCapMultiplier) {
      verdict = "stop"; reasons.push("전환당 비용 상한 초과");
    } else if (c.inflowCost !== null && avgInflowCost > 0 && c.inflowCost > avgInflowCost * settings.inflowCostOutlierMultiplier && c.realRoas < 1) {
      verdict = "stop"; reasons.push(`유입단가 평균 ${settings.inflowCostOutlierMultiplier}배↑`);
    } else if (c.revenue === 0 && c.inflowCost !== null && avgInflowCost > 0 && c.inflowCost <= avgInflowCost * settings.watchInflowCostRatio) {
      verdict = "watch"; reasons.push(`아직 매출은 없지만 유입단가 ₩${fmtNum(c.inflowCost)}로 평균의 절반 이하 — 관찰 필요`);
    } else {
      if (c.composite >= settings.thFocus) verdict = "focus";
      else if (c.composite >= settings.thKeep) verdict = "keep";
      else if (c.composite >= settings.thCut) verdict = "cut";
      else verdict = "stop";
      reasons.push(`종합지수 ${fmtNum(c.composite, 2)} 기준 판정`);
    }
    c.verdict = verdict;
    c.reason = reasons.join(" · ");
    if (verdict === "stop" && !reasons[0].startsWith("종합지수")) warnings.push({ code: c.code, reason: reasons[0] });
  }

  const ranked = [...creatives].sort((a, b) => (b.composite - a.composite) || (b.revenue - a.revenue));
  ranked.forEach((c, i) => { c.rank = i + 1; });

  const totalPool = creatives.reduce((s, c) => s + c.spend, 0);
  for (const c of ranked) {
    if (c.rank > settings.rankCap || c.verdict === "stop") {
      c.proposedBudget = 0;
      c.note = c.verdict === "stop" ? "판정 중단" : `순위 ${c.rank}위 — 상한(${settings.rankCap}개) 밖, 선택과 집중 컷`;
    } else if (c.verdict === "focus") {
      c.proposedBudget = c.spend * (1 + settings.focusIncreasePct / 100);
      c.note = `집중 배정 +${settings.focusIncreasePct}%`;
    } else if (c.verdict === "cut") {
      c.proposedBudget = c.spend * (1 - settings.cutReducePct / 100);
      c.note = `축소 −${settings.cutReducePct}%`;
    } else {
      c.proposedBudget = c.spend;
      c.note = "유지/관찰 — 동결";
    }
  }
  const sumBefore = ranked.reduce((s, c) => s + c.proposedBudget, 0);
  let remaining = totalPool - sumBefore;
  if (remaining > 0) {
    let pool = ranked.filter((c) => c.rank <= settings.rankCap && c.verdict !== "stop")
      .map((c) => ({ c, cap: c.spend * (1 + settings.perCreativeCapPct / 100) }));
    for (let iter = 0; iter < 4 && remaining > 1 && pool.length; iter++) {
      const wSum = pool.reduce((s, p) => s + Math.max(p.c.composite, 0.01), 0);
      const next = [];
      let distributed = 0;
      for (const p of pool) {
        const share = remaining * (Math.max(p.c.composite, 0.01) / wSum);
        const room = Math.max(p.cap - p.c.proposedBudget, 0);
        const add = Math.min(share, room);
        p.c.proposedBudget += add;
        distributed += add;
        if (room - add > 1) next.push(p);
      }
      remaining -= distributed;
      pool = next;
    }
  }
  for (const c of ranked) {
    c.delta = c.proposedBudget - c.spend;
    c.deltaPct = c.spend > 0 ? (c.delta / c.spend) * 100 : (c.proposedBudget > 0 ? 100 : 0);
  }

  const totalSpend = creatives.reduce((s, c) => s + c.spend, 0);
  const totalRevenue = creatives.reduce((s, c) => s + c.revenue, 0);
  const totalInflow = creatives.reduce((s, c) => s + c.inflow, 0);
  const totalPurchaseValue = creatives.reduce((s, c) => { const m = metaAgg.get(c.code); return s + (m ? m.purchaseValue : 0); }, 0);
  const realRoasTotal = totalSpend > 0 ? totalRevenue / totalSpend : 0;
  const inflowCostAvgTotal = totalInflow > 0 ? totalSpend / totalInflow : 0;
  const overReportRatio = totalRevenue > 0 ? totalPurchaseValue / totalRevenue : null;
  const nextBudgetTotal = ranked.reduce((s, c) => s + c.proposedBudget, 0);

  const badgeCounts = { focus: 0, keep: 0, watch: 0, cut: 0, stop: 0 };
  for (const c of creatives) badgeCounts[c.verdict] = (badgeCounts[c.verdict] || 0) + 1;

  const mediaMap = new Map();
  for (const c of creatives) mediaMap.set(c.media, (mediaMap.get(c.media) || 0) + c.spend);
  const mediaBreakdown = [...mediaMap.entries()].map(([name, value]) => ({ name, value }));

  const topSpendChart = [...creatives].sort((a, b) => b.spend - a.spend).slice(0, 12)
    .map((c) => ({ name: c.code, spend: c.spend, revenue: c.revenue }));

  return {
    creatives: ranked, missingSpend, warnings, badgeCounts, mediaBreakdown, topSpendChart,
    kpi: { totalSpend, totalRevenue, realRoasTotal, totalInflow, inflowCostAvgTotal, overReportRatio, nextBudgetTotal, siteAvgRevenuePerInflow },
  };
}

/* ---------------- 렌더링 ---------------- */

function badgeHtml(kind) {
  const s = BADGE_STYLES[kind] || BADGE_STYLES.none;
  return `<span class="badge" style="background:${s.bg};color:${s.fg}">${s.label}</span>`;
}

function barChartSvg(data) {
  const w = 620, rowH = 24, padL = 170, padR = 60, gap = 6;
  const max = Math.max(1, ...data.map((d) => Math.max(d.spend, d.revenue)));
  const h = data.length * (rowH * 2 + gap) + 20;
  const scale = (w - padL - padR) / max;
  let bars = "";
  data.forEach((d, i) => {
    const y0 = 10 + i * (rowH * 2 + gap);
    bars += `<text x="${padL - 8}" y="${y0 + rowH}" text-anchor="end" font-size="10.5" fill="#44506a">${esc(d.name)}</text>`;
    bars += `<rect x="${padL}" y="${y0}" width="${d.spend * scale}" height="${rowH - 4}" fill="#2563eb" rx="3"></rect>`;
    bars += `<text x="${padL + d.spend * scale + 6}" y="${y0 + rowH - 6}" font-size="10" fill="#6b7688">${fmtWon(d.spend)}</text>`;
    bars += `<rect x="${padL}" y="${y0 + rowH}" width="${d.revenue * scale}" height="${rowH - 4}" fill="#059669" rx="3"></rect>`;
    bars += `<text x="${padL + d.revenue * scale + 6}" y="${y0 + rowH * 2 - 6}" font-size="10" fill="#6b7688">${fmtWon(d.revenue)}</text>`;
  });
  return `<svg viewBox="0 0 ${w} ${h}" width="100%" style="max-height:320px">
    <text x="${padL}" y="8" font-size="10" fill="#2563eb">■ 지출</text>
    <text x="${padL + 60}" y="8" font-size="10" fill="#059669">■ 실매출</text>
    ${bars}
  </svg>`;
}

function donutChartSvg(data) {
  const colors = ["#2563eb", "#059669", "#d97706", "#dc2626", "#64748b", "#7c3aed"];
  const total = data.reduce((s, d) => s + d.value, 0) || 1;
  const cx = 110, cy = 110, r = 80, rInner = 46;
  let angle = -90, paths = "", legend = "";
  data.forEach((d, i) => {
    const frac = d.value / total;
    const a0 = angle, a1 = angle + frac * 360;
    angle = a1;
    const large = a1 - a0 > 180 ? 1 : 0;
    const toXY = (a, rad) => [cx + rad * Math.cos((a * Math.PI) / 180), cy + rad * Math.sin((a * Math.PI) / 180)];
    const [x0, y0] = toXY(a0, r), [x1, y1] = toXY(a1, r);
    const [ix1, iy1] = toXY(a1, rInner), [ix0, iy0] = toXY(a0, rInner);
    paths += `<path d="M${x0},${y0} A${r},${r} 0 ${large} 1 ${x1},${y1} L${ix1},${iy1} A${rInner},${rInner} 0 ${large} 0 ${ix0},${iy0} Z" fill="${colors[i % colors.length]}"></path>`;
    legend += `<div style="display:flex;align-items:center;gap:6px;font-size:11.5px;margin-bottom:4px"><span style="width:10px;height:10px;background:${colors[i % colors.length]};border-radius:2px;display:inline-block"></span>${esc(d.name)}</div>`;
  });
  return `<div style="display:flex;align-items:center;gap:20px;justify-content:center">
    <svg viewBox="0 0 220 220" width="180" height="180">${paths}</svg>
    <div>${legend}</div>
  </div>`;
}

function renderDashboard(root, state) {
  const { metaRows, utmRows, settings, dateStart, dateEnd, uploader, savedAt } = state;
  if (!metaRows || !utmRows) {
    root.innerHTML = `<div class="empty-state">아직 데이터가 없습니다. 관리자 페이지에서 데이터를 업로드해주세요.</div>`;
    return;
  }
  const d = computeDashboard(metaRows, utmRows, dateStart, dateEnd, settings);
  root.dataset.sortKey = root.dataset.sortKey || "composite";
  root.dataset.sortDir = root.dataset.sortDir || "desc";

  const sortKey = root.dataset.sortKey, sortDir = root.dataset.sortDir;
  const sorted = [...d.creatives].sort((a, b) => {
    let av = a[sortKey], bv = b[sortKey];
    if (av === null || av === undefined) av = -Infinity;
    if (bv === null || bv === undefined) bv = -Infinity;
    const dir = sortDir === "asc" ? 1 : -1;
    return typeof av === "string" ? dir * av.localeCompare(bv) : dir * (av - bv);
  });

  const kpiHtml = `
    <div class="kpi-grid">
      <div class="kpi"><div class="l">총 지출</div><div class="v">${fmtWon(d.kpi.totalSpend)}</div></div>
      <div class="kpi"><div class="l">UTM 실매출</div><div class="v">${fmtWon(d.kpi.totalRevenue)}</div></div>
      <div class="kpi"><div class="l">실ROAS</div><div class="v" style="color:${d.kpi.realRoasTotal < 1 ? "#dc2626" : "#059669"}">${fmtNum(d.kpi.realRoasTotal, 2)}</div></div>
      <div class="kpi"><div class="l">총 유입 / 유입단가</div><div class="v">${fmtNum(d.kpi.totalInflow)} / ${fmtWon(d.kpi.inflowCostAvgTotal)}</div></div>
      <div class="kpi"><div class="l">매체 과대계상 배율</div><div class="v">${d.kpi.overReportRatio !== null ? fmtNum(d.kpi.overReportRatio, 2) + "배" : "-"}</div><div class="d">매체 보고 전환값 ÷ 실매출</div></div>
      <div class="kpi"><div class="l">차주 제안 예산 합계</div><div class="v">${fmtWon(d.kpi.nextBudgetTotal)}</div><div class="d">현 지출 대비 ${fmtNum((d.kpi.nextBudgetTotal / d.kpi.totalSpend) * 100)}%</div></div>
    </div>`;

  const badgeSummary = Object.entries(d.badgeCounts).filter(([, v]) => v > 0)
    .map(([k, v]) => `${badgeHtml(k)} <b style="font-size:12.5px;margin-right:14px">${v}</b>`).join("");

  const warnHtml = d.warnings.length ? `<div class="warn">⚠ 점검 필요: ${d.warnings.map((w) => `<b>${esc(w.code)}</b>(${esc(w.reason)})`).join(" · ")}</div>` : "";

  const tableRows = sorted.map((c) => `
    <tr>
      <td class="l">${esc(c.code)}</td>
      <td class="l">${esc(c.media)}</td>
      <td>${fmtWon(c.spend)}</td>
      <td>${c.mediaRoas === null ? "-" : fmtNum(c.mediaRoas, 2)}</td>
      <td>${fmtNum(c.inflow)}</td>
      <td>${c.inflowCost === null ? "-" : fmtWon(c.inflowCost)}</td>
      <td>${fmtWon(c.revenuePerInflow)}</td>
      <td style="color:${c.realRoas < 1 ? "#dc2626" : "#059669"};font-weight:600">${fmtNum(c.realRoas, 2)}</td>
      <td>${fmtNum(c.inflowValue, 2)}</td>
      <td style="font-weight:700">${fmtNum(c.composite, 2)}</td>
      <td>${badgeHtml(c.verdict)}</td>
      <td class="why">${esc(c.reason)}</td>
      <td><b>${fmtWon(c.proposedBudget)}</b></td>
    </tr>`).join("");

  const reallocRows = [...d.creatives].sort((a, b) => a.rank - b.rank).map((c) => `
    <tr>
      <td>${c.rank}</td><td class="l">${esc(c.code)}</td><td class="l">${esc(c.media)}</td>
      <td>${badgeHtml(c.verdict)}</td><td>${fmtNum(c.composite, 2)}</td>
      <td>${fmtWon(c.spend)}</td><td><b>${fmtWon(c.proposedBudget)}</b></td>
      <td style="color:${c.delta < 0 ? "#dc2626" : c.delta > 0 ? "#059669" : "#6b7688"};font-weight:600">${fmtWon(c.delta)} (${fmtPct(c.deltaPct)})</td>
      <td class="why">${esc(c.note)}</td>
    </tr>`).join("");

  const missingRows = d.missingSpend.map((r) => `
    <tr><td class="l">${esc(r.code)}</td><td class="l">${esc(r.media)}</td><td>${fmtNum(r.inflow)}</td><td>${fmtNum(r.orders)}</td><td>${fmtWon(r.revenue)}</td></tr>`).join("");

  const sortableTh = (label, key) => `<th data-sort="${key}" class="${sortKey === key ? "sorted-" + sortDir : ""}">${label}</th>`;

  root.innerHTML = `
    <div class="report-meta">분석 기간 ${esc(dateStart)} ~ ${esc(dateEnd)}${uploader ? ` · 업로드: ${esc(uploader)}${savedAt ? ` (${new Date(savedAt).toLocaleString("ko-KR")})` : ""}` : ""}</div>
    ${kpiHtml}
    <div style="margin:14px 0">${badgeSummary}</div>
    ${warnHtml}
    <div class="charts">
      <div class="chartbox"><h3>지출 상위 소재: 지출 vs 실매출</h3>${barChartSvg(d.topSpendChart)}</div>
      <div class="chartbox"><h3>매체별 지출 비중</h3>${donutChartSvg(d.mediaBreakdown)}</div>
    </div>
    <div class="card">
      <h2>소재별 판정표 <span class="hint">열 제목 클릭으로 정렬 · 사이트 평균 유입당매출 ${fmtWon(d.kpi.siteAvgRevenuePerInflow)} 기준</span></h2>
      <div style="overflow-x:auto"><table id="main-table"><thead><tr>
        <th>소재</th><th>매체</th>
        ${sortableTh("지출", "spend")}<th>매체ROAS</th>
        ${sortableTh("유입", "inflow")}${sortableTh("유입단가", "inflowCost")}
        <th>유입당매출</th>${sortableTh("실ROAS", "realRoas")}
        <th>유입가치</th>${sortableTh("종합", "composite")}
        <th>판정</th><th>판정 근거</th>${sortableTh("차주 예산", "proposedBudget")}
      </tr></thead><tbody>${tableRows}</tbody></table></div>
    </div>
    <div class="card">
      <h2>🎯 선택과 집중 — 차주 예산 재배분 제안 <span class="hint">총예산 = 금주 지출 유지 기준 · 상한 ${settings.rankCap}개 / 소재당 최대 +${settings.perCreativeCapPct}%</span></h2>
      <div style="overflow-x:auto"><table><thead><tr>
        <th>순위</th><th>소재</th><th>매체</th><th>판정</th><th>종합</th><th>현 지출</th><th>재배분 예산</th><th>증감</th><th>비고</th>
      </tr></thead><tbody>${reallocRows}</tbody></table></div>
    </div>
    ${d.missingSpend.length ? `<div class="card">
      <h2>지출 데이터 미입력 소재 <span class="hint">UTM 성과는 있으나 매체 지출이 업로드되지 않음</span></h2>
      <table><thead><tr><th>소재</th><th>매체(추정)</th><th>유입</th><th>주문</th><th>매출</th></tr></thead><tbody>${missingRows}</tbody></table>
    </div>` : ""}
  `;

  root.querySelectorAll("th[data-sort]").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.sort;
      if (root.dataset.sortKey === key) root.dataset.sortDir = root.dataset.sortDir === "asc" ? "desc" : "asc";
      else { root.dataset.sortKey = key; root.dataset.sortDir = "desc"; }
      renderDashboard(root, state);
    });
  });
}
