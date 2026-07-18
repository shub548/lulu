/* 룰루피치 광고 판정 대시보드 — 공용 계산/렌더링 로직
   index.html(팀 전체 보기)과 admin.html(업로드) 양쪽에서 그대로 사용합니다. */

var DEFAULT_SETTINGS = {
  // 종합지수 가중치 — "전환이 검증된 소재"에는 실ROAS 비중을 훨씬 높게 둠 (전환=최우선 원칙)
  weightRoasVerified: 0.85,
  weightInflowValueVerified: 0.15,
  thFocus: 2.0,
  thKeep: 1.0,
  thCut: 0.7,
  cpaCapWon: 58000,              // 전환당비용 상한 (원본 리포트에서 역산된 고정값)
  inflowCostOutlierMultiplier: 3, // 유입단가 이상치 = 평균 × N
  noRevenueSpendWon: 100000,      // 무매출 허용선 (고정값, 평균 배수 아님)
  watchInflowCostRatio: 0.5,      // 소액관찰 유입단가 기준 = 평균 × N
  rankCap: 30,
  focusIncreasePct: 25,
  cutReducePct: 45,
  perCreativeCapPct: 50,
  // 신규: 표본 부족 필터 — 이 미만이면 종합지수가 높아도 "집중"을 못 받음
  minInflowForFocus: 20,
  minOrdersForFocus: 2,
  // 신규: 모멘텀(연속 부진 격상) / 피로도 경고
  fatigueDropRatio: 0.4, // 실ROAS가 전기간 대비 40% 이상 하락하면 피로 경고
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
    campaign: r["캠페인 이름"] || null,
    adset: r["광고 세트 이름"] || null,
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
      // 카페24 CSV 앞에 붙는 BOM(﻿) 문자 때문에 "date" 컬럼명이 깨지는 문제 방지
      transformHeader: (h) => h.replace(/^\uFEFF/, "").trim(),
      complete: (res) => {
        resolve(res.data.map((r) => ({
          date: r.date,
          channel: r.channel ? String(r.channel).trim() : null,
          medium: r.medium ? String(r.medium).trim() : null,
          campaign: r.campaign ? String(r.campaign).trim() : null,
          content: r.content ? String(r.content).trim() : null,
          term: r.term ? String(r.term).trim() : null,
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

function computeDashboard(metaRows, utmRows, dateStart, dateEnd, settings, prevMap) {
  const metaInRange = metaRows.filter((r) => inRange(r.date, dateStart, dateEnd));
  const utmInRange = utmRows.filter((r) => inRange(r.date, dateStart, dateEnd));

  const metaAgg = new Map();
  for (const r of metaInRange) {
    const cur = metaAgg.get(r.code) || { spend: 0, purchases: 0, purchaseValue: 0, adName: r.adName, campaign: r.campaign, adset: r.adset };
    cur.spend += r.spend;
    cur.purchases += r.purchases;
    cur.purchaseValue += r.purchaseValue;
    // 최신 행의 이름/캠페인/세트 정보로 갱신 (기간 내 마지막 값 사용)
    cur.adName = r.adName || cur.adName;
    cur.campaign = r.campaign || cur.campaign;
    cur.adset = r.adset || cur.adset;
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
    const verified = revenue > 0; // 전환(매출)이 실제로 확인된 소재인지
    // 전환이 검증된 소재는 실ROAS 비중을 훨씬 높게, 미검증 소재는 유입가치로만 판단
    const composite = verified
      ? realRoas * settings.weightRoasVerified + inflowValue * settings.weightInflowValueVerified
      : inflowValue;
    const cpa = orders > 0 ? spend / orders : null;
    creatives.push({
      code, adName: m.adName, campaign: m.campaign, adset: m.adset, media: classifyMedia(code),
      spend, mediaRoas, inflow, inflowCost, revenuePerInflow, revenue, realRoas, inflowValue,
      composite, orders, cpa, verified,
    });
  }

  const spendCreatives = creatives.filter((c) => c.spend > 0);
  const avgInflowCost = avg(spendCreatives.filter((c) => c.inflowCost !== null).map((c) => c.inflowCost));

  const warnings = [];
  const fatigueWarnings = [];
  for (const c of creatives) {
    const reasons = [];
    let verdict;
    let sampleLimited = false;

    if (c.inflow === 0) {
      verdict = "stop"; reasons.push("유입 0 — UTM 세팅 점검");
    } else if (c.revenue === 0 && c.spend > settings.noRevenueSpendWon) {
      verdict = "cut";
      reasons.push(`실매출 0인데 지출 ${fmtWon(c.spend)} — 무매출 허용선(${fmtWon(settings.noRevenueSpendWon)}) 초과. −${settings.cutReducePct}% 후 차주 재평가`);
    } else if (c.revenue === 0 && c.inflowCost !== null && avgInflowCost > 0 && c.inflowCost <= avgInflowCost * settings.watchInflowCostRatio) {
      verdict = "watch";
      reasons.push(`아직 매출은 없지만 유입단가 ${fmtWon(c.inflowCost)}로 평균의 절반 이하 — 관찰 필요`);
    } else if (c.cpa !== null && c.cpa > settings.cpaCapWon) {
      verdict = "stop";
      reasons.push(`전환 1건 비용 ${fmtWon(c.cpa)} — 상한(${fmtWon(settings.cpaCapWon)}) 초과`);
    } else if (c.inflowCost !== null && avgInflowCost > 0 && c.inflowCost > avgInflowCost * settings.inflowCostOutlierMultiplier && c.realRoas < 1) {
      verdict = "stop";
      reasons.push(`유입단가 ${fmtWon(c.inflowCost)} — 전체 평균(${fmtWon(avgInflowCost)})의 ${settings.inflowCostOutlierMultiplier}배 초과, 매출로도 회수 못함`);
    } else {
      if (c.composite >= settings.thFocus) verdict = "focus";
      else if (c.composite >= settings.thKeep) verdict = "keep";
      else if (c.composite >= settings.thCut) verdict = "cut";
      else verdict = "stop";
      reasons.push(`종합지수 ${fmtNum(c.composite, 2)} 기준 판정${c.verified ? "" : " (미검증 · 유입가치 기준)"}`);

      // 표본 부족 필터: 유입/주문이 너무 적으면 "집중"은 못 받음 (우연한 대박 방지)
      if (verdict === "focus" && (c.inflow < settings.minInflowForFocus || c.orders < settings.minOrdersForFocus)) {
        verdict = "keep";
        sampleLimited = true;
        reasons.push(`표본 부족(유입 ${c.inflow}/${settings.minInflowForFocus}, 주문 ${c.orders}/${settings.minOrdersForFocus}) — 집중 보류, 유지로 조정`);
      }
      // 미검증(매출 미발생) 소재는 유입가치가 아무리 높아도 "집중"까지는 못 감
      if (verdict === "focus" && !c.verified) {
        verdict = "keep";
        reasons.push("미검증(매출 미발생) — 유입가치만으로는 집중 불가, 유지로 조정");
      }
    }

    // 모멘텀: 지난 기간에도 부진했는데 이번에도 부진하면 한 단계 격상 (2주 연속 부진 시 중단)
    if (prevMap && prevMap.has(c.code)) {
      const prev = prevMap.get(c.code);
      if (verdict === "cut" && (prev.verdict === "cut" || prev.verdict === "stop")) {
        reasons.push(`지난 기간(${prev.verdict === "cut" ? "축소" : "중단"})에 이어 이번 기간도 부진 — 중단으로 격상`);
        verdict = "stop";
      } else if (prev.verdict === "watch" && c.revenue > 0) {
        reasons.push("지난 기간 관찰 소재 — 이번 기간 첫 매출 발생 확인");
      }
      // 피로도 경고: 실ROAS가 전기간 대비 크게 하락 (판정 자체는 바꾸지 않고 경고만)
      if (prev.realRoas > 0 && c.realRoas < prev.realRoas * (1 - settings.fatigueDropRatio) && (verdict === "focus" || verdict === "keep")) {
        fatigueWarnings.push({ code: c.code, from: prev.realRoas, to: c.realRoas });
      }
    }

    c.verdict = verdict;
    c.sampleLimited = sampleLimited;
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
      c.note = `집중 배정 +${settings.focusIncreasePct}% (검증된 전환)`;
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
    // 검증된 전환(집중 등급)을 최우선으로 배분하고, 그 다음에야 유지 등급으로 넘어감
    // (전환=최우선 원칙 반영: 미검증/유지 소재보다 실매출 확인된 소재를 우대)
    const tiers = [
      ranked.filter((c) => c.rank <= settings.rankCap && c.verdict === "focus"),
      ranked.filter((c) => c.rank <= settings.rankCap && c.verdict === "keep"),
    ];
    for (const tier of tiers) {
      if (remaining <= 1) break;
      let pool = tier.map((c) => ({ c, cap: c.spend * (1 + settings.perCreativeCapPct / 100) }));
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

  // 다음 기간 모멘텀/피로도 판단에 쓰일 수 있도록 코드별 판정 요약 맵도 함께 반환
  const verdictMap = new Map(creatives.map((c) => [c.code, { verdict: c.verdict, realRoas: c.realRoas }]));

  return {
    creatives: ranked, missingSpend, warnings, fatigueWarnings, badgeCounts, mediaBreakdown, topSpendChart, verdictMap,
    kpi: { totalSpend, totalRevenue, realRoasTotal, totalInflow, inflowCostAvgTotal, overReportRatio, nextBudgetTotal, siteAvgRevenuePerInflow },
  };
}

function rowKeyMeta(r) { return r.code + "|" + r.date; }

// 메타(지출) 데이터 병합: 날짜가 항상 명확한 일자 단위라 코드+날짜 키로 정확히 매칭/교체 가능
function mergeRows(existingRows, newRows, keyFn) {
  const map = new Map();
  for (const r of (existingRows || [])) map.set(keyFn(r), r);
  for (const r of (newRows || [])) map.set(keyFn(r), r);
  return [...map.values()];
}

// 날짜 값이 "2026-07-01" 같은 단일 날짜인지, "2026-07-01 ~ 2026-07-18" 같은 범위인지 파싱해서 [시작,끝]을 반환
function parseDateSpan(dateStr) {
  if (!dateStr) return [null, null];
  const parts = String(dateStr).split("~").map((s) => s.trim());
  if (parts.length === 2 && parts[0] && parts[1]) return [parts[0], parts[1]];
  return [dateStr, dateStr];
}
function spansOverlap(s1, e1, s2, e2) {
  if (!s1 || !e1 || !s2 || !e2) return false;
  return s1 <= e2 && e1 >= s2;
}

// UTM 데이터 전용 병합: 날짜 표기 형식(일자별 vs 기간집계)이 다른 파일끼리 겹치는 기간을 올릴 때
// 같은 매출이 이중으로 쌓이는 걸 막기 위해, 새로 올라온 파일이 커버하는 기간과 겹치는 기존 행은
// 전부 제거하고 새 데이터로 통째로 교체합니다 (행 단위 키 매칭 대신 "기간 단위 교체").
function mergeUtmRowsByPeriod(existingRows, newRows) {
  if (!newRows || !newRows.length) return existingRows || [];
  let newStart = null, newEnd = null;
  for (const r of newRows) {
    const [s, e] = parseDateSpan(r.date);
    if (!s || !e) continue;
    if (newStart === null || s < newStart) newStart = s;
    if (newEnd === null || e > newEnd) newEnd = e;
  }
  const survivors = (existingRows || []).filter((r) => {
    const [s, e] = parseDateSpan(r.date);
    return !spansOverlap(s, e, newStart, newEnd);
  });
  return [...survivors, ...newRows];
}

function shiftDate(dateStr, days) {
  const d = new Date(dateStr + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}
function daysBetween(a, b) {
  return Math.round((new Date(b + "T00:00:00Z") - new Date(a + "T00:00:00Z")) / 86400000) + 1;
}
function previousPeriod(dateStart, dateEnd) {
  const len = daysBetween(dateStart, dateEnd);
  return [shiftDate(dateStart, -len), shiftDate(dateStart, -1)];
}

// 일자별 지출/실매출/유입 추이 (전체 매체 합산, ad-attributed 매출 기준)
function computeDailyTrend(metaRows, utmRows, dateStart, dateEnd) {
  const codeLike = /^\d{6,}/;
  const byDate = new Map();
  const ensure = (date) => {
    if (!byDate.has(date)) byDate.set(date, { date, spend: 0, revenue: 0, inflow: 0 });
    return byDate.get(date);
  };
  for (const r of metaRows) {
    if (!inRange(r.date, dateStart, dateEnd)) continue;
    ensure(r.date).spend += r.spend;
  }
  for (const r of utmRows) {
    if (!inRange(r.date, dateStart, dateEnd)) continue;
    if (r.channel && codeLike.test(r.channel)) {
      const e = ensure(r.date);
      e.revenue += r.orderAmount;
      e.inflow += r.inflowCount;
    }
  }
  return [...byDate.values()].sort((a, b) => a.date.localeCompare(b.date));
}

function lineChartSvg(trend) {
  if (!trend.length) return '<div style="padding:30px;text-align:center;color:var(--sub);font-size:12.5px">해당 기간에 표시할 일자별 데이터가 없습니다</div>';
  const w = 640, h = 220, padL = 60, padR = 20, padT = 16, padB = 26;
  const maxV = Math.max(1, ...trend.map((t) => Math.max(t.spend, t.revenue)));
  const xStep = (w - padL - padR) / Math.max(1, trend.length - 1);
  const yScale = (v) => padT + (h - padT - padB) * (1 - v / maxV);
  const xAt = (i) => padL + i * xStep;
  const pathOf = (key) => trend.map((t, i) => `${i === 0 ? "M" : "L"}${xAt(i).toFixed(1)},${yScale(t[key]).toFixed(1)}`).join(" ");
  const spendPath = pathOf("spend");
  const revPath = pathOf("revenue");
  const dots = (key, color) => trend.map((t, i) => `<circle cx="${xAt(i).toFixed(1)}" cy="${yScale(t[key]).toFixed(1)}" r="2.6" fill="${color}"><title>${t.date}: ${fmtWon(t[key])}</title></circle>`).join("");
  const step = Math.ceil(trend.length / 8) || 1;
  const xLabels = trend.map((t, i) => (i % step === 0 ? `<text x="${xAt(i).toFixed(1)}" y="${h - 6}" font-size="9.5" fill="#6b7688" text-anchor="middle">${t.date.slice(5)}</text>` : "")).join("");
  const gridY = [0, 0.25, 0.5, 0.75, 1].map((f) => {
    const y = padT + (h - padT - padB) * (1 - f);
    return `<line x1="${padL}" x2="${w - padR}" y1="${y}" y2="${y}" stroke="#eef1f6" stroke-width="1"></line><text x="${padL - 6}" y="${y + 3}" font-size="9" fill="#94a3b8" text-anchor="end">${(maxV * f / 10000).toFixed(0)}만</text>`;
  }).join("");
  return `<svg viewBox="0 0 ${w} ${h}" width="100%" style="max-height:240px">
    ${gridY}
    <path d="${spendPath}" fill="none" stroke="#2563eb" stroke-width="2"></path>
    <path d="${revPath}" fill="none" stroke="#059669" stroke-width="2"></path>
    ${dots("spend", "#2563eb")}${dots("revenue", "#059669")}
    ${xLabels}
  </svg>
  <div style="text-align:center;font-size:11px;margin-top:4px">
    <span style="color:#2563eb">■ 지출</span> &nbsp; <span style="color:#059669">■ 실매출(UTM)</span>
  </div>`;
}

function deltaHtml(cur, prev, opts) {
  opts = opts || {};
  if (prev === null || prev === undefined || prev === 0) return "";
  const diff = cur - prev;
  const pct = (diff / Math.abs(prev)) * 100;
  const up = diff > 0;
  const goodIsUp = opts.goodIsUp !== false;
  const color = diff === 0 ? "#6b7688" : (up === goodIsUp ? "#059669" : "#dc2626");
  const arrow = diff === 0 ? "–" : (up ? "▲" : "▼");
  return `<div style="font-size:11px;color:${color};margin-top:3px">${arrow} ${Math.abs(pct).toFixed(1)}% vs 이전 기간</div>`;
}

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

function creativeLabelHtml(c) {
  const campaign = c.campaign ? esc(c.campaign) : "";
  const adset = c.adset ? esc(c.adset) : "";
  const name = c.adName ? esc(c.adName) : esc(c.code);
  return `<div style="line-height:1.4">
    ${campaign || adset ? `<div style="font-size:10px;color:#94a3b8">${campaign}${campaign && adset ? " › " : ""}${adset}</div>` : ""}
    <div style="font-weight:600">${name}</div>
  </div>`;
}

function renderDashboard(root, state, view) {
  const { metaRows, utmRows, settings, uploader, savedAt } = state;
  if (!metaRows || !utmRows) {
    root.innerHTML = `<div class="empty-state">아직 데이터가 없습니다. 관리자 페이지에서 데이터를 업로드해주세요.</div>`;
    return;
  }
  view = view || {};
  const dateStart = view.dateStart || state.dateStart;
  const dateEnd = view.dateEnd || state.dateEnd;
  const mediaFilter = view.mediaFilter || null;   // Set|null
  const verdictFilter = view.verdictFilter || null; // Set|null
  const compare = !!view.compare;

  // 모멘텀/피로도 판단용 — "이전 동일기간과 비교" 표시 여부와 무관하게 항상 계산합니다.
  const [pStart, pEnd] = previousPeriod(dateStart, dateEnd);
  const prevDashboard = computeDashboard(metaRows, utmRows, pStart, pEnd, settings);
  const d = computeDashboard(metaRows, utmRows, dateStart, dateEnd, settings, prevDashboard.verdictMap);
  const trend = computeDailyTrend(metaRows, utmRows, dateStart, dateEnd);

  const cmp = compare ? { range: [pStart, pEnd], d: prevDashboard } : null;

  root.dataset.sortKey = root.dataset.sortKey || "composite";
  root.dataset.sortDir = root.dataset.sortDir || "desc";
  const sortKey = root.dataset.sortKey, sortDir = root.dataset.sortDir;

  let visibleCreatives = d.creatives;
  if (mediaFilter && mediaFilter.size) visibleCreatives = visibleCreatives.filter((c) => mediaFilter.has(c.media));
  if (verdictFilter && verdictFilter.size) visibleCreatives = visibleCreatives.filter((c) => verdictFilter.has(c.verdict));

  const sorted = [...visibleCreatives].sort((a, b) => {
    let av = a[sortKey], bv = b[sortKey];
    if (av === null || av === undefined) av = -Infinity;
    if (bv === null || bv === undefined) bv = -Infinity;
    const dir = sortDir === "asc" ? 1 : -1;
    return typeof av === "string" ? dir * av.localeCompare(bv) : dir * (av - bv);
  });

  const kpiHtml = `
    <div class="kpi-grid">
      <div class="kpi"><div class="l">총 지출</div><div class="v">${fmtWon(d.kpi.totalSpend)}</div>${cmp ? deltaHtml(d.kpi.totalSpend, cmp.d.kpi.totalSpend, { goodIsUp: false }) : ""}</div>
      <div class="kpi"><div class="l">실매출(UTM)</div><div class="v">${fmtWon(d.kpi.totalRevenue)}</div>${cmp ? deltaHtml(d.kpi.totalRevenue, cmp.d.kpi.totalRevenue) : ""}</div>
      <div class="kpi"><div class="l">실ROAS</div><div class="v" style="color:${d.kpi.realRoasTotal < 1 ? "#dc2626" : "#059669"}">${fmtNum(d.kpi.realRoasTotal, 2)}</div>${cmp ? deltaHtml(d.kpi.realRoasTotal, cmp.d.kpi.realRoasTotal) : ""}</div>
      <div class="kpi"><div class="l">총 유입 / 유입단가</div><div class="v">${fmtNum(d.kpi.totalInflow)} / ${fmtWon(d.kpi.inflowCostAvgTotal)}</div>${cmp ? deltaHtml(d.kpi.totalInflow, cmp.d.kpi.totalInflow) : ""}</div>
      <div class="kpi"><div class="l">매체 과대계상 배율</div><div class="v">${d.kpi.overReportRatio !== null ? fmtNum(d.kpi.overReportRatio, 2) + "배" : "-"}</div><div class="d">매체 보고 전환값 ÷ 실매출</div></div>
      <div class="kpi"><div class="l">차주 제안 예산 합계</div><div class="v">${fmtWon(d.kpi.nextBudgetTotal)}</div><div class="d">현 지출 대비 ${fmtNum((d.kpi.nextBudgetTotal / d.kpi.totalSpend) * 100)}%</div></div>
    </div>`;

  const badgeSummary = Object.entries(d.badgeCounts).filter(([, v]) => v > 0)
    .map(([k, v]) => `${badgeHtml(k)} <b style="font-size:12.5px;margin-right:14px">${v}</b>`).join("");

  const warnHtml = d.warnings.length ? `<div class="warn">⚠ 점검 필요: ${d.warnings.map((w) => `<b>${esc(w.code)}</b>(${esc(w.reason)})`).join(" · ")}</div>` : "";
  const fatigueHtml = d.fatigueWarnings.length ? `<div class="warn" style="background:#fef2f2;border-color:#fca5a5;color:#991b1b">🔥 실ROAS 급락(소재 피로 의심): ${d.fatigueWarnings.map((w) => `<b>${esc(w.code)}</b>(${fmtNum(w.from, 2)}→${fmtNum(w.to, 2)})`).join(" · ")}</div>` : "";

  const tableRows = sorted.map((c) => `
    <tr>
      <td class="l">${creativeLabelHtml(c)}</td>
      <td class="l">${esc(c.media)}</td>
      <td>${fmtWon(c.spend)}</td>
      <td>${fmtWon(c.revenue)}</td>
      <td>${c.mediaRoas === null ? "-" : fmtNum(c.mediaRoas, 2)}</td>
      <td>${fmtNum(c.inflow)}</td>
      <td>${c.inflowCost === null ? "-" : fmtWon(c.inflowCost)}</td>
      <td>${fmtWon(c.revenuePerInflow)}</td>
      <td style="color:${c.realRoas < 1 ? "#dc2626" : "#059669"};font-weight:600">${fmtNum(c.realRoas, 2)}</td>
      <td>${fmtNum(c.inflowValue, 2)}</td>
      <td style="font-weight:700">${fmtNum(c.composite, 2)}</td>
      <td>${badgeHtml(c.verdict)}${c.verified ? "" : ' <span style="font-size:10px;color:#94a3b8">미검증</span>'}${c.sampleLimited ? ' <span style="font-size:10px;color:#a16207">표본부족</span>' : ""}</td>
      <td class="why">${esc(c.reason)}</td>
      <td><b>${fmtWon(c.proposedBudget)}</b></td>
    </tr>`).join("");

  const reallocRows = [...visibleCreatives].sort((a, b) => a.rank - b.rank).map((c) => `
    <tr>
      <td>${c.rank}</td><td class="l">${creativeLabelHtml(c)}</td><td class="l">${esc(c.media)}</td>
      <td>${badgeHtml(c.verdict)}</td><td>${fmtNum(c.composite, 2)}</td>
      <td>${fmtWon(c.spend)}</td><td><b>${fmtWon(c.proposedBudget)}</b></td>
      <td style="color:${c.delta < 0 ? "#dc2626" : c.delta > 0 ? "#059669" : "#6b7688"};font-weight:600">${fmtWon(c.delta)} (${fmtPct(c.deltaPct)})</td>
      <td class="why">${esc(c.note)}</td>
    </tr>`).join("");

  const missingRows = d.missingSpend.map((r) => `
    <tr><td class="l">${esc(r.code)}</td><td class="l">${esc(r.media)}</td><td>${fmtNum(r.inflow)}</td><td>${fmtNum(r.orders)}</td><td>${fmtWon(r.revenue)}</td></tr>`).join("");

  const sortableTh = (label, key) => `<th data-sort="${key}" class="${sortKey === key ? "sorted-" + sortDir : ""}">${label}</th>`;
  const filterNote = (mediaFilter && mediaFilter.size) || (verdictFilter && verdictFilter.size)
    ? `<div class="note" style="margin-bottom:10px">필터 적용 중 — 아래 표는 ${visibleCreatives.length}개 소재만 표시 (순위·차주예산은 전체 ${d.creatives.length}개 기준 계산값)</div>` : "";

  root.innerHTML = `
    <div class="report-meta">분석 기간 ${esc(dateStart)} ~ ${esc(dateEnd)}${cmp ? ` · 비교 기간 ${esc(cmp.range[0])} ~ ${esc(cmp.range[1])}` : ""}${uploader ? ` · 최근 업로드: ${esc(uploader)}${savedAt ? ` (${new Date(savedAt).toLocaleString("ko-KR")})` : ""}` : ""}</div>
    ${kpiHtml}
    <div style="margin:14px 0">${badgeSummary}</div>
    ${warnHtml}
    ${fatigueHtml}
    <div class="card">
      <h2>일자별 추이 <span class="hint">선택 기간 내 매체 지출 vs UTM 실매출(귀속 매출)</span></h2>
      ${lineChartSvg(trend)}
    </div>
    <div class="charts">
      <div class="chartbox"><h3>지출 상위 소재: 지출 vs 실매출</h3>${barChartSvg(d.topSpendChart)}</div>
      <div class="chartbox"><h3>매체별 지출 비중</h3>${donutChartSvg(d.mediaBreakdown)}</div>
    </div>
    <div class="card">
      <h2>소재별 판정표 <span class="hint">열 제목 클릭으로 정렬 · 사이트 평균 유입당매출 ${fmtWon(d.kpi.siteAvgRevenuePerInflow)} 기준 · "미검증"=아직 매출 미발생(유입가치로만 판단, 집중 불가)</span></h2>
      ${filterNote}
      <div style="overflow-x:auto"><table id="main-table"><thead><tr>
        <th>캠페인 › 세트 › 소재</th><th>매체</th>
        ${sortableTh("지출", "spend")}${sortableTh("실매출", "revenue")}<th>매체ROAS</th>
        ${sortableTh("유입", "inflow")}${sortableTh("유입단가", "inflowCost")}
        <th>유입당매출</th>${sortableTh("실ROAS", "realRoas")}
        <th>유입가치</th>${sortableTh("종합", "composite")}
        <th>판정</th><th>판정 근거</th>${sortableTh("차주 예산", "proposedBudget")}
      </tr></thead><tbody>${tableRows}</tbody></table></div>
    </div>
    <div class="card">
      <h2>🎯 선택과 집중 — 차주 예산 재배분 제안 <span class="hint">총예산 = 선택 기간 지출 유지 기준 · 상한 ${settings.rankCap}개 / 소재당 최대 +${settings.perCreativeCapPct}% · 검증된 전환(집중) 우선 배분</span></h2>
      <div style="overflow-x:auto"><table><thead><tr>
        <th>순위</th><th>캠페인 › 세트 › 소재</th><th>매체</th><th>판정</th><th>종합</th><th>현 지출</th><th>재배분 예산</th><th>증감</th><th>비고</th>
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
      renderDashboard(root, state, view);
    });
  });
}
