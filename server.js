const fs = require("node:fs");
const path = require("node:path");
const express = require("express");
const AdmZip = require("adm-zip");
const ExcelJS = require("exceljs");
const { DatabaseSync } = require("node:sqlite");

loadEnv();

const app = express();
const port = Number(process.env.PORT || 3000);
const dartKey = process.env.OPENDART_API_KEY || "";
const openaiKey = process.env.OPENAI_API_KEY || "";
const openaiModel = process.env.OPENAI_MODEL || "gpt-4.1-mini";
const dataDir = path.join(__dirname, "data");
const dbPath = path.join(dataDir, "bit-analysis.sqlite");
const companyTemplatePath = path.join(__dirname, "templates", "company-analysis-template.xlsx");
const DART_BASE = "https://opendart.fss.or.kr/api";
const OPENAI_RESPONSES_URL = "https://api.openai.com/v1/responses";

fs.mkdirSync(dataDir, { recursive: true });

const db = new DatabaseSync(dbPath);
initDb();
seedCorpMaster();

app.use(express.json({ limit: "1mb" }));
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});
app.use(express.static(__dirname));

app.get("/api/health", (req, res) => {
  res.json({
    ok: true,
    dart_key_configured: Boolean(dartKey),
    openai_key_configured: Boolean(openaiKey),
    openai_model: openaiModel,
    db_path: dbPath,
    corp_count: db.prepare("SELECT COUNT(*) AS count FROM corp_master").get().count
  });
});

app.get("/api/company/refresh-master", async (req, res) => {
  try {
    const result = await refreshCorpMaster({ force: true });
    res.json(result);
  } catch (error) {
    res.status(400).json({ ok: false, error: error.message });
  }
});

app.get("/api/company/resolve", async (req, res) => {
  const query = String(req.query.q || "").trim();
  const initialMatches = resolveCompanies(query);
  if (initialMatches.length || !query) {
    return res.json({
      query,
      source: dartKey ? "sqlite-seed" : "sqlite-seed",
      matches: initialMatches
    });
  }

  try {
    await refreshCorpMaster({ force: false });
  } catch (error) {
    // Seed data still works without a DART key or network.
  }
  const matches = resolveCompanies(query);
  res.json({
    query,
    source: dartKey ? "sqlite+opendart-cache" : "sqlite-seed",
    matches
  });
});

app.get("/api/dart/financials", async (req, res) => {
  const corpCode = String(req.query.corp_code || "").trim();
  const baseYear = Number(req.query.bsns_year || req.query.base_year || new Date().getFullYear() - 1);
  const years = clamp(Number(req.query.years || 5), 1, 7);
  const reportCode = String(req.query.report_code || "11011");
  const fsDiv = String(req.query.fs_div || "CFS").toUpperCase();

  if (!dartKey) {
    return res.status(400).json({ ok: false, error: "OPENDART_API_KEY is not configured. Create .env from .env.example." });
  }
  if (!/^\d{8}$/.test(corpCode)) {
    return res.status(400).json({ ok: false, error: "corp_code must be 8 digits." });
  }

  try {
    res.json(await buildFinancialsResponse({ corpCode, baseYear, years, reportCode, fsDiv }));
  } catch (error) {
    res.status(502).json({ ok: false, error: error.message });
  }
});

app.get("/api/company/market-data", async (req, res) => {
  const corpCode = String(req.query.corp_code || "").trim();
  const stockCode = String(req.query.stock_code || "").trim();
  const baseYear = Number(req.query.bsns_year || req.query.base_year || new Date().getFullYear() - 1);
  const reportCode = String(req.query.report_code || "11011");

  if (!stockCode) {
    return res.status(400).json({ ok: false, error: "stock_code is required." });
  }

  const company =
    (corpCode && db.prepare("SELECT * FROM corp_master WHERE corp_code = ?").get(corpCode)) ||
    db.prepare("SELECT * FROM corp_master WHERE stock_code = ? LIMIT 1").get(stockCode) ||
    { corp_code: corpCode, corp_name: stockCode, stock_code: stockCode, market: "LISTED" };

  try {
    res.json({
      ok: true,
      market_data: await buildMarketData({
        company: publicCorp(company),
        corpCode: corpCode || company.corp_code || "",
        baseYear,
        reportCode
      })
    });
  } catch (error) {
    res.status(502).json({ ok: false, error: error.message });
  }
});

app.get("/api/export/company-workbook", async (req, res) => {
  const corpCode = String(req.query.corp_code || "").trim();
  const baseYear = Number(req.query.bsns_year || req.query.base_year || new Date().getFullYear() - 1);
  const years = clamp(Number(req.query.years || 5), 1, 7);
  const reportCode = String(req.query.report_code || "11011");
  const fsDiv = String(req.query.fs_div || "CFS").toUpperCase();

  if (!dartKey) {
    return res.status(400).json({ ok: false, error: "OPENDART_API_KEY is not configured. Create .env from .env.example." });
  }
  if (!/^\d{8}$/.test(corpCode)) {
    return res.status(400).json({ ok: false, error: "corp_code must be 8 digits." });
  }

  try {
    const financials = await buildFinancialsResponse({ corpCode, baseYear, years, reportCode, fsDiv });
    const buffer = await buildCompanyWorkbook(financials);
    const safeName = encodeURIComponent(`${financials.company.name || corpCode}_기업분석.xlsx`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${safeName}`);
    res.send(buffer);
  } catch (error) {
    res.status(502).json({ ok: false, error: error.message });
  }
});

app.post("/api/analysis/company", async (req, res) => {
  const financials = req.body && req.body.financials;
  const reportType = String(req.body && (req.body.report_type || req.body.reportType) || "상세 분석").trim();
  if (!financials || !Array.isArray(financials.periods)) {
    return res.status(400).json({ ok: false, error: "financials.periods is required." });
  }
  const latest = [...financials.periods].reverse().find(period => period.metrics && Object.keys(period.metrics).length);
  if (!latest) {
    return res.json({ ok: true, generated_by: "local-fallback", analysis: buildCompanyAnalysisFallback(financials, reportType) });
  }

  const fallback = () => buildCompanyAnalysisFallback(financials, reportType);
  if (!openaiKey) {
    return res.json({
      ok: true,
      generated_by: "local-fallback",
      warning: "OPENAI_API_KEY is not configured. Using numeric fallback.",
      analysis: fallback()
    });
  }

  try {
    const analysis = await generateCompanyAnalysis({ financials, reportType });
    res.json({ ok: true, generated_by: "openai", model: openaiModel, analysis });
  } catch (error) {
    res.json({
      ok: true,
      generated_by: "local-fallback",
      warning: `OpenAI generation failed: ${error.message}`,
      analysis: fallback()
    });
  }
});

app.post("/api/analysis/industry", async (req, res) => {
  const body = req.body || {};
  const industry = String(body.industry || "").trim();
  const scope = String(body.scope || "").trim();
  const period = String(body.period || "").trim();
  const reportType = String(body.report_type || body.reportType || "").trim();

  if (!industry) return res.status(400).json({ ok: false, error: "industry is required." });
  if (!scope) return res.status(400).json({ ok: false, error: "scope is required." });
  if (!period) return res.status(400).json({ ok: false, error: "period is required." });
  if (!reportType) return res.status(400).json({ ok: false, error: "report_type is required." });

  const fallback = () => buildIndustryFallback({ industry, scope, period, reportType });
  if (!openaiKey) {
    return res.json({
      ok: true,
      generated_by: "local-fallback",
      warning: "OPENAI_API_KEY is not configured. Using local analyst-style fallback.",
      analysis: fallback()
    });
  }

  try {
    const analysis = await generateIndustryAnalysis({ industry, scope, period, reportType });
    res.json({ ok: true, generated_by: "openai", model: openaiModel, analysis });
  } catch (error) {
    res.json({
      ok: true,
      generated_by: "local-fallback",
      warning: `OpenAI generation failed: ${error.message}`,
      analysis: fallback()
    });
  }
});

const server = app.listen(port, () => {
  console.log(`BIT Analysis server listening on http://127.0.0.1:${port}`);
  if (!dartKey) console.log("OPENDART_API_KEY is not configured yet. Copy .env.example to .env and add your key.");
  if (!openaiKey) console.log("OPENAI_API_KEY is not configured yet. Industry AI generation will use local fallback.");
});
server.ref();

function loadEnv() {
  const envPath = path.join(__dirname, ".env");
  if (!fs.existsSync(envPath)) return;
  const lines = fs.readFileSync(envPath, "utf8").split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).trim();
    const value = trimmed.slice(eq + 1).trim().replace(/^["']|["']$/g, "");
    if (!process.env[key]) process.env[key] = value;
  }
}

async function fetchWithTimeout(url, options = {}, timeoutMs = 10000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
}

async function generateIndustryAnalysis({ industry, scope, period, reportType }) {
  const industryContext = resolveIndustryContext(industry);
  const analysisIndustry = industryContext.analysisIndustry;
  const isDeep = isDeepIndustryReport(reportType);
  const payload = {
    model: openaiModel,
    instructions: [
      "You are a Korean consulting research analyst writing a detailed industry analysis report.",
      "Analyze the industry only. Do not mix company analysis, DART disclosures, company revenue, EBIT, EBITDA, balance-sheet metrics, or valuation metrics.",
      "Write the actual report, not advice, instructions, interview coaching, resume coaching, or methodology.",
      "Avoid vague guidance such as '봐야 합니다', '확인해야 합니다', '설명하면 좋습니다'. Instead state what is happening in the industry and why it matters.",
      "Each long string field should be 2-4 complete Korean sentences with concrete sector context, not a one-line placeholder.",
      "For value chain rows, write in plain business language. Do not use abstract phrases like '대체재가 적거나 병목을 가진 공급자' without naming the actual actor group or bottleneck type.",
      "If the user's industry input is narrow, colloquial, or ambiguous, reframe it into a researchable parent industry and then analyze the original input as a segment inside that parent industry.",
      "When reframing, state the reframed scope in market_scope and executive_summary, but do not turn it into user advice. The report should read as an industry report.",
      "Use both Korean and English keyword context when thinking about sparse categories, especially for consumer goods, IP goods, and niche manufacturing segments.",
      "Use an answer-first, causal, detailed Korean business-report tone.",
      "Market trend must be market-size level data, not an index like 100/120/80. Use a real unit such as '십억 달러', '조원', '백만 대', or another industry-appropriate unit.",
      "Do not fabricate exact market sizes, stock returns, financial statements, URLs, or news dates.",
      "This request does not include live web-browsing evidence. Therefore market_trend.series must mark every point as is_estimated true unless the user supplied verified source data in the prompt.",
      "If exact data is uncertain, use broad estimated market-size figures, explain the uncertainty in data_quality_note and source_reference, and write sources as source categories or source candidates rather than fake citations.",
      "For sources, do not invent report titles or URLs. Use source candidate categories such as '통계청/산업통상자원부/협회 통계 검증 필요' or '글로벌 리서치 기관 검증 필요' and leave url empty when no specific verified URL was supplied.",
      "Keep the same report structure regardless of report type; 요약 Brief means compact but still substantive, not shallow.",
      "Return only valid JSON matching the requested schema."
    ].join("\n"),
    input: [
      `사용자 입력 산업명: ${industryContext.originalIndustry}`,
      `분석 기준 산업명: ${analysisIndustry}`,
      `상위 산업/자료 탐색 범위: ${industryContext.parentIndustry}`,
      `세부 세그먼트: ${industryContext.segments.join(", ")}`,
      `한글/영문 리서치 키워드: ${industryContext.keywords.join(", ")}`,
      `범위 재정의 메모: ${industryContext.rationale}`,
      `분석 범위: ${scope}`,
      `분석 기간: ${period}`,
      `리포트 유형: ${reportType}`,
      "웹앱에 바로 렌더링할 수 있도록 다음 흐름을 지켜줘: 산업 개요, 시장 규모/성장 추이 graph, 밸류체인, Revenue/Cost 구조, 핵심 변수, 산업적 시사점.",
      isDeep
        ? "이번 요청은 '심층 분석'이다. 증권사 산업 리포트처럼 핵심 결론, 주요 이슈 플로우, 세부 섹터별 판단, 추적 지표, 시나리오, 리스크를 deep_dive에 충분한 분량으로 작성해줘."
        : "이번 요청은 기본 리포트다. deep_dive는 생성하지 말고 기본 산출물 필드만 충실하게 작성해줘.",
      "나쁜 예: '수요 변화와 비용 구조를 함께 봐야 합니다.'",
      "좋은 예: '자동차 산업은 판매대수보다 파워트레인 믹스와 지역별 가격 정책이 수익성을 가르는 국면입니다. 미국과 유럽에서는 전기차 재고와 인센티브 부담이 커지는 반면, 하이브리드는 연비 수요와 낮은 보조금 의존도 때문에 완성차 업체의 마진 방어 수단으로 기능하고 있습니다.'",
      "산업적 시사점은 '무엇을 해야 하는가'가 아니라 '이 산업은 어떤 구조로 움직이며, 어디서 이익과 리스크가 발생하는가'에 답해줘."
    ].join("\n"),
    max_output_tokens: isDeep ? 9000 : 6500,
    text: {
      format: {
        type: "json_schema",
        name: "industry_analysis",
        strict: true,
        schema: industryAnalysisSchema({ deep: isDeep })
      }
    }
  };

  const response = await fetchWithTimeout(OPENAI_RESPONSES_URL, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${openaiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  }, 60000);

  const data = await response.json().catch(() => null);
  if (!response.ok) {
    const message = data && data.error && data.error.message ? data.error.message : `HTTP ${response.status}`;
    throw new Error(message);
  }

  const text = extractResponseText(data);
  if (!text) throw new Error("OpenAI response did not contain output text.");
  const parsed = JSON.parse(text);
  return normalizeIndustryAnalysis(parsed, { industry: analysisIndustry, scope, period, reportType, industryContext });
}

function isDeepIndustryReport(reportType) {
  return /심층|deep/i.test(String(reportType || ""));
}

function resolveIndustryContext(rawIndustry) {
  const originalIndustry = String(rawIndustry || "").trim();
  const normalized = normalizeSearchText(originalIndustry.replace(/산업$/g, ""));
  const makeContext = ({
    patterns,
    analysisIndustry,
    parentIndustry,
    segments,
    keywords,
    marketUnit,
    marketSeries,
    rationale,
    valueChain = null,
    revenueSources = null,
    costDrivers = null
  }) => ({
    patterns,
    analysisIndustry,
    parentIndustry,
    segments,
    keywords,
    marketUnit,
    marketSeries,
    rationale,
    valueChain,
    revenueSources,
    costDrivers
  });
  const rules = [
    makeContext({
      patterns: [/소비재/, /생활소비재/, /일반소비재/, /consumergoods/, /consumergoods/, /consumerstaples/, /fmcg/, /cpg/],
      analysisIndustry: "소비재 및 FMCG",
      parentIndustry: "식품, 생활용품, 퍼스널케어, 화장품, 패션잡화, 가정용품, 온라인/오프라인 유통 채널",
      segments: ["식품/음료", "생활용품·퍼스널케어", "화장품·뷰티", "패션·잡화", "D2C/온라인 브랜드", "대형 유통·편의점·이커머스"],
      keywords: ["소비재 산업", "FMCG 시장", "생활용품 시장", "consumer goods market", "FMCG market", "consumer staples", "CPG market", "retail consumer products"],
      marketUnit: "국내 소비재 소매 판매액 추정치, 조원",
      marketSeries: [490, 512, 545, 575, 603],
      rationale: "소비재는 품목 범위가 넓어 단일 제품 시장으로 해석하면 답변이 흐려집니다. 따라서 식품, 생활용품, 퍼스널케어, 패션잡화, 유통 채널을 포함한 FMCG/소매 소비재 상위 시장으로 재정의해 분석합니다.",
      valueChain: [
        { stage: "Brand / Product Planning", participants: "브랜드사, ODM/OEM 기획팀, 카테고리 매니저, 상품기획 조직", role: "소비자 니즈, 가격대, 패키지, 브랜드 포지션을 제품으로 설계합니다.", revenue_model: "브랜드 제품 판매, PB/ODM 납품, 라이선스 상품 매출이 발생합니다.", cost_structure: "제품 개발비, 디자인·패키징, 샘플링, 브랜드 마케팅비가 주요 비용입니다.", margin_power: "강한 브랜드와 반복 구매 카테고리를 가진 사업자가 유통사보다 높은 가격 방어력을 확보합니다.", report_implication: "소비재의 profit pool은 단순 제조보다 브랜드, 카테고리 점유율, 반복 구매 빈도에 집중됩니다." },
        { stage: "Manufacturing / Sourcing", participants: "원재료 공급자, OEM/ODM 제조사, 포장재 업체, 품질 인증 기관", role: "제품을 안정적으로 생산하고 원가, 품질, 납기, 규제 대응을 관리합니다.", revenue_model: "생산 단가, 장기 납품 계약, ODM 개발 생산에서 매출이 발생합니다.", cost_structure: "원재료, 포장재, 인건비, 에너지, 물류, 품질관리 비용이 핵심입니다.", margin_power: "레시피, 원료 조달, 인증, 소량 다품종 생산 역량이 있으면 단순 OEM보다 협상력이 올라갑니다.", report_implication: "원가 상승기에는 제조 경쟁력보다 브랜드의 가격 전가력과 유통 협상력이 수익성을 가릅니다." },
        { stage: "Retail / Channel", participants: "대형마트, 편의점, 이커머스, H&B 스토어, 자사몰, 라이브커머스", role: "제품 노출, 프로모션, 가격, 재고 회전, 소비자 데이터 축적을 담당합니다.", revenue_model: "유통 마진, 입점 수수료, 광고 노출, PB 매출, 배송/멤버십 수익이 발생합니다.", cost_structure: "물류비, 재고 보관비, 판촉비, 수수료, 반품 비용이 부담입니다.", margin_power: "소비자 접점과 데이터를 가진 채널은 브랜드의 판매 조건과 프로모션 강도를 좌우합니다.", report_implication: "소비재에서는 브랜드와 채널의 협상력 변화가 매출보다 마진을 더 직접적으로 흔듭니다." },
        { stage: "Consumer Demand", participants: "가계 소비자, MZ/시니어 세그먼트, 가격 민감 소비층, 프리미엄 소비층", role: "구매 빈도, 객단가, 브랜드 충성도, 가격 민감도를 통해 산업 수요를 결정합니다.", revenue_model: "반복 구매, 시즌성 구매, 프리미엄 업셀링, 구독/정기배송이 매출을 만듭니다.", cost_structure: "소비자 입장에서는 물가, 소득, 금리, 배송비, 프로모션 조건이 구매 결정 비용입니다.", margin_power: "필수재는 물량 방어가 좋고, 선택재는 브랜드·트렌드·소득 민감도가 큽니다.", report_implication: "소비재 분석은 총소비보다 필수재/선택재 믹스와 가격 전가 후 수요 이탈 여부를 분리해야 합니다." }
      ],
      revenueSources: [
        { source: "Repeat Purchase", mechanism: "식품, 생활용품, 퍼스널케어처럼 구매 주기가 짧은 품목이 안정적 매출을 만듭니다.", sensitivity: "물가와 소득이 압박될 때도 필수재는 물량 방어가 가능하지만, 저가/프로모션 전환이 마진을 낮출 수 있습니다." },
        { source: "Premium / Brand Mix", mechanism: "프리미엄 라인, 기능성 제품, 브랜드 충성도가 높은 카테고리에서 ASP와 마진이 올라갑니다.", sensitivity: "소비 심리가 약해지면 프리미엄 전환 속도가 둔화되고 중저가 PB와 경쟁이 커집니다." },
        { source: "Channel Expansion", mechanism: "이커머스, 편의점, H&B, 자사몰, 라이브커머스가 신규 고객 접점과 반복 구매 데이터를 만듭니다.", sensitivity: "채널 수수료, 배송비, 광고비가 상승하면 매출 성장에도 영업 레버리지가 약해질 수 있습니다." }
      ],
      costDrivers: [
        { cost: "Raw material / packaging", mechanism: "곡물, 유지, 화학 원료, 포장재, 환율이 제조 원가를 압박합니다.", risk: "브랜드가 가격을 올려도 소비자가 PB나 저가 대체재로 이동하면 마진 회복이 제한됩니다." },
        { cost: "Promotion / channel fee", mechanism: "입점 수수료, 광고비, 판촉비, 라이브커머스 비용이 판매관리비를 구성합니다.", risk: "채널 의존도가 높으면 매출 성장을 위해 판촉비를 계속 써야 하는 구조가 됩니다." },
        { cost: "Inventory / logistics", mechanism: "SKU 확장, 유통기한, 반품, 배송비가 재고와 물류 비용을 키웁니다.", risk: "수요 예측 실패 시 할인 판매와 재고평가손실이 동시에 발생합니다." }
      ]
    }),
    makeContext({
      patterns: [/인형/, /봉제완구/, /완구인형/, /plush/, /doll/, /stuffedtoy/, /키덜트/, /캐릭터굿즈/],
      analysisIndustry: "완구 및 캐릭터 IP 굿즈",
      parentIndustry: "완구, 봉제완구, 캐릭터 라이선싱 상품, 키덜트 수집형 굿즈",
      segments: ["유아동 완구", "봉제완구·플러시 토이", "패션/수집형 인형", "캐릭터 IP 라이선싱 상품", "키덜트·팬덤 굿즈"],
      keywords: ["완구 산업", "봉제완구", "인형 시장", "캐릭터 상품", "키덜트 시장", "toy market", "plush toy market", "doll market", "licensed character merchandise", "collectibles market"],
      marketUnit: "글로벌 완구/캐릭터 상품 시장 규모 추정치, 십억 달러",
      marketSeries: [103, 108, 114, 121, 129],
      rationale: "인형은 독립 산업 통계가 제한적이므로 완구·캐릭터 IP 굿즈를 상위 시장으로 두고, 봉제완구와 수집형 인형을 세부 세그먼트로 분석합니다.",
      valueChain: [
        { stage: "IP / Design", participants: "캐릭터 IP 보유사, 디자인 스튜디오, 라이선싱 에이전시", role: "캐릭터 세계관, 디자인 원형, 라이선스 권리를 제공하며 제품 차별화의 출발점이 됩니다.", revenue_model: "라이선스 로열티, 디자인 용역, 공동기획 수수료에서 매출이 발생합니다.", cost_structure: "IP 개발비, 디자인 인력비, 마케팅비, 법무·계약 비용이 주요 부담입니다.", margin_power: "인지도 높은 IP와 팬덤을 가진 권리자는 제조사보다 높은 협상력을 확보합니다.", report_implication: "인형 세그먼트의 이익은 단순 제조보다 IP 권리와 팬덤 전환력에 더 크게 좌우됩니다." },
        { stage: "Materials / OEM", participants: "원단·충전재 공급자, 봉제 OEM/ODM, 안전 인증 기관", role: "봉제완구와 인형을 실제 제품으로 전환하고 품질, 안전성, 납기, 원가를 결정합니다.", revenue_model: "OEM 생산 단가, ODM 기획 생산, 대량 납품 계약으로 매출이 발생합니다.", cost_structure: "원단·충전재, 인건비, 물류비, 안전 인증과 품질관리 비용이 핵심입니다.", margin_power: "대체 생산지가 많으면 마진은 낮지만, 소량 다품종·고품질 봉제 역량이나 인증 경험이 있으면 협상력이 올라갑니다.", report_implication: "제조 단계는 규모보다 품질 안정성과 빠른 SKU 전환 능력이 수익성을 가릅니다." },
        { stage: "Brand / Merchandising", participants: "완구 브랜드, 캐릭터 상품사, 팬덤 굿즈 기획사, 유통 PB", role: "상품 콘셉트, 가격대, 한정판 전략, 패키징을 설계해 수요를 제품 매출로 전환합니다.", revenue_model: "완제품 판매, 한정판 드롭, 콜라보 상품, 굿즈 번들 매출이 발생합니다.", cost_structure: "재고 부담, 마케팅비, 라이선스 비용, 반품·폐기 비용이 주요 리스크입니다.", margin_power: "브랜드 팬덤과 희소성 설계가 가능하면 가격 할인 없이도 판매가 유지됩니다.", report_implication: "수집형 인형은 완구보다 패션·팬덤 소비재에 가까워 재고 회전과 희소성 관리가 중요합니다." },
        { stage: "Retail / Platform", participants: "온라인몰, 대형마트, 팬덤 플랫폼, 팝업스토어, 역직구 채널", role: "제품을 소비자에게 노출하고 구매 전환, 예약 판매, 재판매 가격 형성에 영향을 줍니다.", revenue_model: "유통 마진, 플랫폼 수수료, 광고·노출 상품, 팝업 매출이 발생합니다.", cost_structure: "물류비, 판매수수료, 재고 보관비, 오프라인 운영비가 비용으로 작동합니다.", margin_power: "팬덤 접근성과 예약·한정 판매 데이터를 가진 채널이 높은 교섭력을 가집니다.", report_implication: "채널은 단순 판매처가 아니라 수요 예측과 가격 방어를 좌우하는 데이터 접점입니다." }
      ],
      revenueSources: [
        { source: "Licensed IP Goods", mechanism: "캐릭터·애니메이션·게임 IP를 인형과 굿즈로 전환해 완제품 판매와 로열티 매출을 만듭니다.", sensitivity: "IP 흥행, 신작 공개, 팬덤 규모, 라이선스 계약 조건에 민감합니다." },
        { source: "Collectible / Limited Edition", mechanism: "한정판, 시즌 상품, 콜라보 제품을 통해 일반 완구보다 높은 ASP와 반복 구매를 만듭니다.", sensitivity: "희소성 설계가 약하거나 재고가 과도하면 할인 판매로 마진이 훼손됩니다." },
        { source: "Mass Toy Retail", mechanism: "유아동 완구 채널과 온라인몰에서 물량 기반 매출이 발생합니다.", sensitivity: "출산율, 소비 경기, 대형 유통 채널의 가격 정책에 민감합니다." }
      ],
      costDrivers: [
        { cost: "License / IP cost", mechanism: "인기 캐릭터를 사용할수록 선급금과 로열티 부담이 커집니다.", risk: "IP 흥행이 기대보다 낮으면 고정성 라이선스 비용이 재고 손실로 전이됩니다." },
        { cost: "Textile / labor / safety", mechanism: "원단, 충전재, 봉제 인건비, KC·CE 등 안전 인증 비용이 제조 원가를 구성합니다.", risk: "저가 제품 경쟁이 심하면 원가 상승분을 소비자가격에 전가하기 어렵습니다." },
        { cost: "Inventory / channel", mechanism: "시즌성 상품과 캐릭터 유행 상품은 재고 보관, 반품, 폐기 비용이 발생합니다.", risk: "수요 예측 실패 시 가격 할인과 재고평가손실이 동시에 나타납니다." }
      ]
    }),
    makeContext({
      patterns: [/문구/, /팬시/, /스티커/, /다꾸/, /stationery/, /fancygoods/],
      analysisIndustry: "문구 및 팬시 굿즈",
      parentIndustry: "문구, 팬시상품, 캐릭터 굿즈, 오피스·학습용품",
      segments: ["학습·사무 문구", "다이어리/꾸미기 용품", "캐릭터 팬시상품", "온라인 소량 브랜드", "B2B 오피스 소모품"],
      keywords: ["문구 산업", "팬시 산업", "다꾸 시장", "캐릭터 문구", "stationery market", "fancy goods market", "character stationery"],
      marketUnit: "문구 및 팬시상품 시장 규모 추정치, 조원 또는 십억 달러",
      marketSeries: [7.8, 8.1, 8.5, 8.9, 9.4],
      rationale: "문구·팬시류는 단일 품목보다 학습/사무 수요, 캐릭터 IP, 온라인 소량 브랜드가 결합된 상위 시장으로 분석해야 자료 탐색이 안정적입니다."
    }),
    makeContext({
      patterns: [/피규어/, /collectible/, /collectibles/, /트레이딩카드/, /tcg/, /보드게임/],
      analysisIndustry: "수집형 취미 및 테이블탑 게임",
      parentIndustry: "키덜트 취미재, 피규어, TCG, 보드게임, 팬덤 커머스",
      segments: ["피규어/스태츄", "트레이딩 카드", "보드게임", "팬덤 한정판", "중고·리셀 플랫폼"],
      keywords: ["키덜트 시장", "피규어 시장", "트레이딩 카드 게임", "보드게임 산업", "collectibles market", "trading card game market", "tabletop games market"],
      marketUnit: "수집형 취미재 시장 규모 추정치, 십억 달러",
      marketSeries: [28, 31, 35, 39, 44],
      rationale: "수집형 취미재는 완구 통계만으로 설명하기 어렵고 팬덤, IP, 리셀, 한정판 가격 형성까지 함께 봐야 합니다."
    }),
    makeContext({
      patterns: [/향수/, /디퓨저/, /캔들/, /fragrance/, /perfume/, /homefragrance/],
      analysisIndustry: "프래그런스 및 홈센트",
      parentIndustry: "퍼스널 프래그런스, 홈 프래그런스, 니치 향수, 라이프스타일 소비재",
      segments: ["니치 향수", "매스 프래그런스", "디퓨저/캔들", "라이프스타일 편집숍", "온라인 D2C 브랜드"],
      keywords: ["향수 시장", "니치 향수", "디퓨저 시장", "fragrance market", "perfume market", "home fragrance market", "niche fragrance"],
      marketUnit: "프래그런스 시장 규모 추정치, 십억 달러",
      marketSeries: [56, 59, 63, 67, 72],
      rationale: "향 관련 품목은 화장품, 라이프스타일, 홈센트로 자료가 흩어져 있어 프래그런스 상위 시장과 세부 채널을 함께 잡아야 합니다."
    }),
    makeContext({
      patterns: [/캠핑/, /등산/, /아웃도어/, /outdoor/, /camping/, /hiking/],
      analysisIndustry: "아웃도어 및 캠핑용품",
      parentIndustry: "아웃도어 의류, 캠핑 장비, 레저용품, 체험형 여가 소비",
      segments: ["아웃도어 의류", "캠핑 장비", "등산/트레킹 용품", "렌탈·중고 장비", "레저 플랫폼"],
      keywords: ["캠핑 산업", "아웃도어 시장", "등산용품", "outdoor gear market", "camping equipment market", "hiking gear market"],
      marketUnit: "아웃도어/캠핑용품 시장 규모 추정치, 십억 달러",
      marketSeries: [78, 82, 86, 90, 95],
      rationale: "캠핑·등산 같은 취미 카테고리는 장비, 의류, 플랫폼, 중고/렌탈 시장이 함께 움직이므로 아웃도어 레저용품으로 확장합니다."
    }),
    makeContext({
      patterns: [/반려/, /펫/, /강아지/, /고양이/, /petcare/, /petfood/, /pet/],
      analysisIndustry: "펫케어 및 반려동물 용품/서비스",
      parentIndustry: "펫푸드, 반려동물 용품, 동물병원, 펫보험, 돌봄 서비스",
      segments: ["펫푸드", "용품/장난감", "동물병원/의료", "펫보험", "미용·호텔·돌봄"],
      keywords: ["펫케어 산업", "반려동물 시장", "펫푸드 시장", "pet care market", "pet food market", "pet services market"],
      marketUnit: "펫케어 시장 규모 추정치, 십억 달러",
      marketSeries: [255, 270, 287, 306, 327],
      rationale: "반려동물 관련 품목은 단일 제품보다 식품, 의료, 보험, 서비스가 결합된 펫케어 생태계로 분석하는 편이 적절합니다."
    }),
    makeContext({
      patterns: [/디저트/, /베이커리/, /빵/, /케이크/, /커피/, /카페/, /bakery/, /dessert/, /coffee/],
      analysisIndustry: "디저트·베이커리 및 카페 F&B",
      parentIndustry: "외식, 베이커리, 카페, 디저트 전문점, 간편식/프랜차이즈",
      segments: ["프랜차이즈 카페", "베이커리", "디저트 전문점", "RTD/간편 디저트", "배달·온라인 예약"],
      keywords: ["디저트 시장", "베이커리 산업", "카페 시장", "bakery market", "dessert market", "coffee shop market", "foodservice market"],
      marketUnit: "디저트/베이커리 F&B 시장 규모 추정치, 조원 또는 십억 달러",
      marketSeries: [42, 45, 48, 52, 56],
      rationale: "디저트·카페류는 품목 하나보다 외식 채널, 프랜차이즈, 원재료, 배달/예약 플랫폼이 결합된 F&B 시장으로 분석해야 합니다."
    })
  ];

  const matched = rules.find(rule => rule.patterns.some(pattern => pattern.test(normalized)));
  if (matched) {
    return {
      originalIndustry,
      analysisIndustry: matched.analysisIndustry,
      parentIndustry: matched.parentIndustry,
      segments: matched.segments,
      keywords: matched.keywords,
      marketUnit: matched.marketUnit,
      marketSeries: matched.marketSeries,
      rationale: matched.rationale,
      valueChain: matched.valueChain,
      revenueSources: matched.revenueSources,
      costDrivers: matched.costDrivers,
      isReframed: true
    };
  }

  const cleanIndustry = originalIndustry.replace(/\s*산업\s*$/g, "").trim() || originalIndustry;
  const looksNiche = cleanIndustry.length <= 8 || !/(자동차|반도체|조선|철강|은행|보험|증권|게임|바이오|제약|화장품|유통|물류|항공|호텔|건설|부동산|통신|미디어|엔터|배터리|에너지|정유|화학|식품|의류|패션|교육|의료|방산|로봇|소프트웨어|클라우드|플랫폼)/.test(cleanIndustry);
  return {
    originalIndustry,
    analysisIndustry: cleanIndustry,
    parentIndustry: looksNiche ? `${cleanIndustry}의 상위 산업, 인접 소비재/서비스 시장, 관련 밸류체인` : cleanIndustry,
    segments: looksNiche
      ? [`${cleanIndustry} 직접 시장`, `${cleanIndustry}가 속한 상위 카테고리`, `${cleanIndustry} 대체재/인접재`, `${cleanIndustry} 유통·플랫폼`, `${cleanIndustry} 공급망`]
      : [`${cleanIndustry} 최종 수요`, `${cleanIndustry} 공급망`, `${cleanIndustry} 생산/운영`, `${cleanIndustry} 채널/서비스`],
    keywords: [cleanIndustry, `${cleanIndustry} 산업`, `${cleanIndustry} 시장`, `${cleanIndustry} value chain`, `${cleanIndustry} market`, `${cleanIndustry} industry`, `${cleanIndustry} adjacent market`, `${cleanIndustry} segment`],
    marketUnit: "시장 규모 추정치, 산업별 적정 단위",
    marketSeries: null,
    rationale: looksNiche
      ? "입력 산업명이 세부 품목 또는 좁은 세그먼트일 가능성이 있어, 상위 산업과 인접 시장을 함께 추론해 분석합니다."
      : "입력 산업명을 그대로 분석 기준으로 사용합니다.",
    valueChain: null,
    revenueSources: null,
    costDrivers: null,
    isReframed: looksNiche
  };
}

function normalizeSearchText(value) {
  return String(value || "").toLowerCase().replace(/[\s.,·_\-\/]/g, "");
}

function industryAnalysisSchema(options = {}) {
  const deep = Boolean(options.deep);
  return {
    type: "object",
    additionalProperties: false,
    required: [
      "title",
      "meta",
      "industry_overview",
      "market_trend",
      "value_chain",
      "revenue_cost_structure",
      "key_variables",
      "report_implications",
      ...(deep ? ["deep_dive"] : []),
      "sources"
    ],
    properties: {
      title: { type: "string" },
      meta: { type: "array", minItems: 3, maxItems: 5, items: { type: "string" } },
      industry_overview: {
        type: "object",
        additionalProperties: false,
        required: ["executive_summary", "definition", "market_scope", "current_state", "demand_side", "supply_side", "structural_change"],
        properties: {
          executive_summary: { type: "string" },
          definition: { type: "string" },
          market_scope: { type: "string" },
          current_state: { type: "string" },
          demand_side: { type: "string" },
          supply_side: { type: "string" },
          structural_change: { type: "string" }
        }
      },
      market_trend: {
        type: "object",
        additionalProperties: false,
        required: ["unit", "series", "growth_comment", "interpretation", "data_quality_note"],
        properties: {
          unit: { type: "string" },
          series: {
            type: "array",
            minItems: 5,
            maxItems: 5,
            items: {
              type: "object",
              additionalProperties: false,
              required: ["year", "value", "is_estimated", "note", "source_reference"],
              properties: {
                year: { type: "string" },
                value: { type: "number" },
                is_estimated: { type: "boolean" },
                note: { type: "string" },
                source_reference: { type: "string" }
              }
            }
          },
          growth_comment: { type: "string" },
          data_quality_note: { type: "string" },
          interpretation: { type: "string" }
        }
      },
      value_chain: {
        type: "array",
        minItems: 4,
        maxItems: 6,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["stage", "participants", "role", "revenue_model", "cost_structure", "margin_power", "report_implication"],
          properties: {
            stage: { type: "string" },
            participants: { type: "string" },
            role: { type: "string" },
            revenue_model: { type: "string" },
            cost_structure: { type: "string" },
            margin_power: { type: "string" },
            report_implication: { type: "string" }
          }
        }
      },
      revenue_cost_structure: {
        type: "object",
        additionalProperties: false,
        required: ["revenue_sources", "cost_drivers", "profit_pool_insight"],
        properties: {
          revenue_sources: {
            type: "array",
            minItems: 3,
            maxItems: 5,
            items: {
              type: "object",
              additionalProperties: false,
              required: ["source", "mechanism", "sensitivity"],
              properties: {
                source: { type: "string" },
                mechanism: { type: "string" },
                sensitivity: { type: "string" }
              }
            }
          },
          cost_drivers: {
            type: "array",
            minItems: 3,
            maxItems: 5,
            items: {
              type: "object",
              additionalProperties: false,
              required: ["cost", "mechanism", "risk"],
              properties: {
                cost: { type: "string" },
                mechanism: { type: "string" },
                risk: { type: "string" }
              }
            }
          },
          profit_pool_insight: { type: "string" }
        }
      },
      key_variables: {
        type: "array",
        minItems: 4,
        maxItems: 6,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["variable", "current_signal", "impact_on_industry", "affected_value_chain", "evidence_to_watch"],
          properties: {
            variable: { type: "string" },
            current_signal: { type: "string" },
            impact_on_industry: { type: "string" },
            affected_value_chain: { type: "string" },
            evidence_to_watch: { type: "string" }
          }
        }
      },
      report_implications: {
        type: "object",
        additionalProperties: false,
        required: ["core_conclusion", "implications", "discussion_points"],
        properties: {
          core_conclusion: { type: "string" },
          implications: { type: "array", minItems: 3, maxItems: 5, items: { type: "string" } },
          discussion_points: { type: "array", minItems: 2, maxItems: 4, items: { type: "string" } }
        }
      },
      ...(deep ? { deep_dive: industryDeepDiveSchema() } : {}),
      sources: {
        type: "array",
        minItems: 2,
        maxItems: 5,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["title", "publisher", "url", "used_for"],
          properties: {
            title: { type: "string" },
            publisher: { type: "string" },
            url: { type: "string" },
            used_for: { type: "string" }
          }
        }
      }
    }
  };
}

function industryDeepDiveSchema() {
  return {
    type: "object",
    additionalProperties: false,
    required: ["headline_thesis", "issue_flow", "subsector_breakdown", "indicator_watch", "scenario_analysis", "risk_factors"],
    properties: {
      headline_thesis: { type: "string" },
      issue_flow: {
        type: "array",
        minItems: 3,
        maxItems: 5,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["issue", "evidence", "industry_impact"],
          properties: {
            issue: { type: "string" },
            evidence: { type: "string" },
            industry_impact: { type: "string" }
          }
        }
      },
      subsector_breakdown: {
        type: "array",
        minItems: 3,
        maxItems: 6,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["subsector", "current_condition", "revenue_logic", "cost_or_bottleneck", "outlook"],
          properties: {
            subsector: { type: "string" },
            current_condition: { type: "string" },
            revenue_logic: { type: "string" },
            cost_or_bottleneck: { type: "string" },
            outlook: { type: "string" }
          }
        }
      },
      indicator_watch: {
        type: "array",
        minItems: 4,
        maxItems: 8,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["indicator", "current_read", "why_it_moves_the_sector"],
          properties: {
            indicator: { type: "string" },
            current_read: { type: "string" },
            why_it_moves_the_sector: { type: "string" }
          }
        }
      },
      scenario_analysis: {
        type: "array",
        minItems: 3,
        maxItems: 3,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["case_name", "conditions", "industry_result"],
          properties: {
            case_name: { type: "string" },
            conditions: { type: "string" },
            industry_result: { type: "string" }
          }
        }
      },
      risk_factors: {
        type: "array",
        minItems: 3,
        maxItems: 6,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["risk", "transmission_path", "early_signal"],
          properties: {
            risk: { type: "string" },
            transmission_path: { type: "string" },
            early_signal: { type: "string" }
          }
        }
      }
    }
  };
}

function extractResponseText(response) {
  if (response && typeof response.output_text === "string") return response.output_text;
  const chunks = [];
  for (const item of response && Array.isArray(response.output) ? response.output : []) {
    for (const content of Array.isArray(item.content) ? item.content : []) {
      if (content.type === "output_text" && content.text) chunks.push(content.text);
    }
  }
  return chunks.join("");
}

function normalizeIndustryAnalysis(analysis, request) {
  const fallback = buildIndustryFallback(request);
  const isDeep = isDeepIndustryReport(request.reportType);
  const normalizedMarketTrend = normalizeMarketTrend(analysis.market_trend, fallback.market_trend);
  const normalized = {
    ...fallback,
    ...analysis,
    title: analysis.title || fallback.title,
    meta: Array.isArray(analysis.meta) && analysis.meta.length ? analysis.meta : fallback.meta,
    industry_overview: {
      ...fallback.industry_overview,
      ...(analysis.industry_overview || {})
    },
    market_trend: {
      ...normalizedMarketTrend
    },
    value_chain: Array.isArray(analysis.value_chain) && analysis.value_chain.length ? analysis.value_chain : fallback.value_chain,
    revenue_cost_structure: {
      ...fallback.revenue_cost_structure,
      ...(analysis.revenue_cost_structure || {})
    },
    key_variables: Array.isArray(analysis.key_variables) && analysis.key_variables.length ? analysis.key_variables : fallback.key_variables,
    report_implications: {
      ...fallback.report_implications,
      ...(analysis.report_implications || {})
    },
    sources: normalizeIndustrySources(analysis.sources, fallback.sources)
  };
  if (isDeep) {
    normalized.deep_dive = {
      ...fallback.deep_dive,
      ...(analysis.deep_dive || {}),
      issue_flow: Array.isArray(analysis.deep_dive && analysis.deep_dive.issue_flow) ? analysis.deep_dive.issue_flow : fallback.deep_dive.issue_flow,
      subsector_breakdown: Array.isArray(analysis.deep_dive && analysis.deep_dive.subsector_breakdown) ? analysis.deep_dive.subsector_breakdown : fallback.deep_dive.subsector_breakdown,
      indicator_watch: Array.isArray(analysis.deep_dive && analysis.deep_dive.indicator_watch) ? analysis.deep_dive.indicator_watch : fallback.deep_dive.indicator_watch,
      scenario_analysis: Array.isArray(analysis.deep_dive && analysis.deep_dive.scenario_analysis) ? analysis.deep_dive.scenario_analysis : fallback.deep_dive.scenario_analysis,
      risk_factors: Array.isArray(analysis.deep_dive && analysis.deep_dive.risk_factors) ? analysis.deep_dive.risk_factors : fallback.deep_dive.risk_factors
    };
  } else {
    delete normalized.deep_dive;
  }
  return normalized;
}

function normalizeMarketTrend(trend, fallback) {
  const unit = String(trend && trend.unit || "").trim();
  const fallbackUnit = String(fallback && fallback.unit || "시장 규모 추정치, 단위 검증 필요").trim();
  return {
    ...fallback,
    ...(trend || {}),
    unit: unit && !/적정 단위|확인|미정|unknown|n\/a/i.test(unit) ? unit : fallbackUnit,
    series: normalizeMarketSeries(trend && trend.series, fallback.series),
    growth_comment: String(trend && trend.growth_comment || fallback.growth_comment || ""),
    interpretation: String(trend && trend.interpretation || fallback.interpretation || ""),
    data_quality_note: ensureEstimatedDataQualityNote(trend && trend.data_quality_note, fallback.data_quality_note)
  };
}

function normalizeMarketSeries(series, fallback) {
  if (!Array.isArray(series) || series.length !== 5) return fallback;
  return series.map((item, index) => {
    const fallbackItem = fallback[index] || {};
    const value = Number(item && item.value);
    return {
      year: String((item && item.year) || fallbackItem.year || ""),
      value: Number.isFinite(value) ? value : Number(fallbackItem.value || 0),
      is_estimated: true,
      note: ensureEstimatedNote((item && item.note) || fallbackItem.note || ""),
      source_reference: ensureEstimatedSourceReference((item && item.source_reference) || fallbackItem.source_reference || "")
    };
  });
}

function ensureEstimatedNote(note) {
  const text = String(note || "").trim();
  if (!text) return "외부 원문 데이터 미연결 상태의 방향성 추정치";
  if (/실제 집계|확정|공식 집계|verified/i.test(text)) return `${text} (서비스 내 원문 검증 전 추정치로 표시)`;
  return text;
}

function ensureEstimatedSourceReference(reference) {
  const text = String(reference || "").trim();
  if (!text) return "원문 출처 검증 필요";
  if (/local fallback|검증 필요|추정|candidate/i.test(text)) return text;
  return `${text} 후보, 원문 검증 필요`;
}

function ensureEstimatedDataQualityNote(note, fallbackNote) {
  const text = String(note || fallbackNote || "").trim();
  const base = text || "정확한 시장 규모는 협회/정부 통계/리서치 원문으로 검증해야 합니다.";
  return /추정|검증|원문|uncertain|estimate/i.test(base)
    ? base
    : `${base} 현재 값은 원문 데이터 미연결 상태의 추정치이므로 별도 검증이 필요합니다.`;
}

function normalizeIndustrySources(sources, fallbackSources) {
  const rows = Array.isArray(sources) && sources.length ? sources : fallbackSources;
  return rows.map(source => {
    const title = String(source && source.title || "출처 후보").trim();
    const publisher = String(source && source.publisher || "검증 필요").trim();
    const usedFor = String(source && source.used_for || "시장 규모와 산업 구조 검증").trim();
    const looksVerifiedUrl = source && typeof source.url === "string" && /^https?:\/\/[^/\s]+\/?$/.test(source.url.trim());
    return {
      title: /검증|후보|candidate/i.test(title) ? title : `${title} 후보`,
      publisher: /검증|후보|candidate/i.test(publisher) ? publisher : `${publisher} 검증 필요`,
      url: looksVerifiedUrl ? source.url.trim() : "",
      used_for: /검증|추정|후보/i.test(usedFor) ? usedFor : `${usedFor} 검증`
    };
  }).slice(0, 5);
}


function normalizeSeries(series, fallback) {
  if (!Array.isArray(series) || series.length !== 5 || series.some(value => !Number.isFinite(Number(value)))) return fallback;
  return series.map(Number);
}

function buildIndustryFallback({ industry, scope, period, reportType, industryContext: providedIndustryContext }) {
  const industryContext = providedIndustryContext || resolveIndustryContext(industry);
  const displayIndustry = industryContext.analysisIndustry || industry;
  const currentYear = new Date().getFullYear();
  const years = Array.from({ length: 5 }, (_, index) => String(currentYear - 4 + index));
  const valueChain = industryContext.valueChain || [
    { stage: "Upstream", participants: "원재료 공급자, 핵심 부품사, 기술/IP 보유 기업, 설비 업체", role: "산업 생산에 필요한 투입 요소와 기술 기반을 제공합니다. 이 구간의 공급 안정성이 낮으면 하위 단계의 생산량과 원가가 동시에 흔들립니다.", revenue_model: "장기 공급계약, 부품 판매, 기술 라이선스, 장비 판매에서 매출이 발생합니다.", cost_structure: "원재료 조달비, 연구개발비, 설비투자, 인증 비용이 주요 부담입니다.", margin_power: "대체 공급자가 적고 고객 인증 장벽이 높은 품목일수록 가격 협상력이 높습니다.", report_implication: "산업 수익성은 최종 수요보다 핵심 투입 요소의 공급 부족 여부에 먼저 반응할 수 있습니다." },
    { stage: "Production / Operation", participants: "제조사, 운영사, 서비스 제공자, 품질·공정 관리 조직", role: "투입 요소를 제품이나 서비스로 전환하고 품질, 수율, 납기, 운영 효율을 관리합니다.", revenue_model: "제품 판매, 프로젝트 수주, 서비스 이용료, 생산량 기반 계약에서 매출이 발생합니다.", cost_structure: "감가상각, 인건비, 에너지 비용, 외주비, 품질보증비 등 고정비와 준고정비 부담이 큽니다.", margin_power: "가동률이 올라갈수록 고정비 흡수 효과가 커지지만, 수요 둔화기에는 손익 훼손이 빠르게 나타납니다.", report_implication: "생산 단계의 핵심은 성장률보다 수율, 가동률, 원가 전가력입니다." },
    { stage: "Channel / Platform", participants: "유통사, 플랫폼, 대형 고객사, B2B 영업망, 애프터마켓 사업자", role: "제품과 서비스를 고객에게 전달하고 반복 구매, 유지보수, 데이터, 고객 관계를 축적합니다.", revenue_model: "유통 마진, 플랫폼 수수료, 구독료, 유지보수, 장기 계약 매출이 발생합니다.", cost_structure: "판매관리비, 물류비, 고객 획득 비용, 플랫폼 운영비가 주요 비용입니다.", margin_power: "고객 전환 비용이 높거나 유통 접점이 제한된 경우 채널 단계가 높은 협상력을 확보합니다.", report_implication: "채널을 장악한 기업은 제조 마진이 낮아져도 반복 매출과 고객 데이터를 통해 산업 내 이익을 방어할 수 있습니다." },
    { stage: "End Demand", participants: "최종 소비자, 기업 고객, 정부·공공기관, 글로벌 바이어", role: "산업의 전체 수요 규모와 투자 사이클을 결정합니다.", revenue_model: "최종 구매, 교체 수요, 기업 투자, 공공 조달, 보조금 기반 수요에서 매출 기회가 발생합니다.", cost_structure: "최종 고객 단계에서는 가격 민감도, 금융비용, 규제 비용, 보조금 변화가 구매 결정을 좌우합니다.", margin_power: "수요가 강해도 최종 고객의 가격 민감도가 높으면 상위 밸류체인의 가격 인상 여지는 제한됩니다.", report_implication: "최종 수요 확대가 산업 전체의 이익 증가로 연결되려면 가격 인상과 비용 전가가 동시에 가능해야 합니다." }
  ];
  return {
    title: `${displayIndustry} 산업분석`,
    meta: industryContext.isReframed ? ["산업분석", scope, period, reportType, `범위 재정의: ${industryContext.originalIndustry}`] : ["산업분석", scope, period, reportType],
    industry_overview: {
      executive_summary: `${displayIndustry} 산업은 최종 수요, 공급망, 가격 결정 구조, 규제 변화가 동시에 작동하는 시장입니다. ${industryContext.rationale} 현재 결과는 입력값을 바탕으로 시장 구조를 먼저 정리한 예비 리포트이며, 원문 리서치 검증을 더하면 시장 규모와 출처 신뢰도를 보강할 수 있습니다.`,
      definition: `${displayIndustry} 산업은 ${industryContext.segments.join(", ")}을 포함해 제품·서비스 생산자, 핵심 투입 요소 공급자, 유통·플랫폼, 최종 고객 수요가 연결된 가치사슬로 구성됩니다. 산업의 매력도는 시장 규모뿐 아니라 가격 결정권, 고정비 부담, 고객 락인, 규제 변화가 어떻게 결합되는지에 따라 달라집니다.`,
      market_scope: `${scope} 범위에서는 ${industryContext.parentIndustry}까지 자료 탐색 범위를 확장해 봅니다. 특히 세부 품목의 직접 통계가 부족한 경우 상위 산업 규모, 인접 세그먼트 성장률, 채널 데이터, 영문 키워드(${industryContext.keywords.slice(0, 4).join(", ")})를 함께 읽어야 시장의 실제 위치가 드러납니다.`,
      current_state: `${period} 기준으로는 수요 성장과 비용 부담이 동시에 존재하는 혼재 국면입니다. ${displayIndustry}은 특정 품목 수요만으로 판단하기보다 상위 카테고리의 소비 흐름, 유통 채널 변화, 공급 원가 변동을 함께 반영해야 합니다.`,
      demand_side: `수요는 ${industryContext.segments.slice(0, 3).join(", ")}의 구매 빈도, 가격대, 소비자 관심도, 기업 고객의 예산 집행에 의해 형성됩니다. 단기 유행이 강해도 반복 구매나 장기 계약 구조가 약하면 실적 가시성은 낮아질 수 있습니다.`,
      supply_side: "공급 측면에서는 핵심 부품, 원재료, 인력, 설비, 기술 인증, 물류, 에너지 비용이 병목으로 작동합니다. 공급 제약이 있는 구간은 가격 협상력을 얻지만, 대체 공급과 증설이 빠르게 진행되면 같은 구간이 중기적으로 마진 압박 요인으로 바뀔 수 있습니다.",
      structural_change: "구조적으로는 단순 물량 성장보다 고부가 제품 믹스, 플랫폼화, 장기계약, 운영 효율화가 산업 이익을 좌우하는 방향으로 이동하고 있습니다. 따라서 산업 분석은 시장 규모 확대와 이익이 남는 가치사슬 구간을 함께 분리해 읽는 구조입니다."
    },
    market_trend: {
      unit: industryContext.marketUnit,
      series: years.map((year, index) => ({
        year,
        value: (industryContext.marketSeries || [10, 11.2, 12.6, 14.3, 16.1])[index],
        is_estimated: true,
        note: industryContext.isReframed ? "세부 품목 직접 통계가 제한적이어서 상위 산업과 인접 시장 기준으로 추정" : "외부 원문 리서치 미연결 상태의 시장 규모 방향성 추정",
        source_reference: "원문 출처 검증 필요"
      })),
      growth_comment: "현재 값은 외부 원문 통계가 직접 연결되지 않은 방향성 추정치이므로 절대값보다 성장 흐름 중심으로 해석해야 합니다.",
      data_quality_note: "정식 리포트에서는 협회 통계, 정부 통계, 증권사 리포트, 글로벌 리서치 기관 자료로 시장 규모 단위를 확정해야 합니다.",
      interpretation: `${period} 관점에서 ${displayIndustry} 산업은 시장 규모가 커지는 흐름과 비용 부담이 함께 나타나는 구조입니다. 따라서 시장 확대가 곧바로 산업 수익성 개선으로 이어지는지, 아니면 특정 가치사슬 구간에만 이익이 집중되는지를 분리해 해석해야 합니다.`
    },
    value_chain: valueChain,
    revenue_cost_structure: {
      revenue_sources: industryContext.revenueSources || [
        { source: "Volume", mechanism: "최종 수요, 교체주기, 고객사 투자 집행이 물량을 만듭니다.", sensitivity: "경기와 고객 예산에 민감합니다." },
        { source: "Price / ASP", mechanism: "공급 부족, 고부가 제품 믹스, 브랜드/기술 우위가 가격을 만듭니다.", sensitivity: "경쟁 심화와 증설 속도에 민감합니다." },
        { source: "Recurring / Contract", mechanism: "장기계약, 구독, 유지보수, 반복 구매가 매출 안정성을 만듭니다.", sensitivity: "고객 락인과 전환 비용에 민감합니다." }
      ],
      cost_drivers: industryContext.costDrivers || [
        { cost: "Input cost", mechanism: "원재료, 핵심 부품, 에너지, 물류비가 원가를 압박합니다.", risk: "가격 전가력이 약하면 매출 증가에도 마진이 훼손됩니다." },
        { cost: "Fixed cost / CAPEX", mechanism: "설비 투자, 감가상각, 인건비, R&D가 고정비 부담을 만듭니다.", risk: "수요 둔화 시 가동률 하락이 이익률을 빠르게 낮춥니다." },
        { cost: "Compliance / Quality", mechanism: "규제 대응, 인증, 품질보증, 안전 비용이 필요합니다.", risk: "인증 지연과 품질 이슈는 출하와 수익성을 동시에 훼손합니다." }
      ],
      profit_pool_insight: `${displayIndustry} 산업의 profit pool은 최종 수요가 가장 큰 곳이 아니라, 공급 병목·고객 락인·가격 전가력을 가진 구간에 남습니다.`
    },
    key_variables: [
      { variable: "최종 수요 성장", current_signal: "수요 자체는 산업 외형을 지지하지만, 가격 경쟁이 심하면 수익성 개선은 제한될 수 있습니다.", impact_on_industry: "수요가 꾸준하면 가동률과 투자 명분은 좋아지지만, 공급 확대가 더 빠르면 마진은 압박받습니다.", affected_value_chain: "End Demand, Production / Operation", evidence_to_watch: "시장 규모, 침투율, 고객사 CAPEX, 재고 수준" },
      { variable: "공급 병목과 증설 속도", current_signal: "단기 병목은 가격 협상력을 만들지만, 중기 증설이 집중되면 공급 과잉 리스크가 커집니다.", impact_on_industry: "병목 구간은 산업 이익을 흡수하고, 과잉 증설 구간은 가격 하락을 유발합니다.", affected_value_chain: "Upstream, Production / Operation", evidence_to_watch: "증설 계획, 리드타임, 가동률, 수주잔고" },
      { variable: "가격 전가력", current_signal: "원가 변동이 큰 산업일수록 가격 전가 여부가 매출 성장보다 중요해집니다.", impact_on_industry: "가격 전가가 가능한 기업은 비용 상승기에도 마진을 방어하고, 그렇지 못한 기업은 외형 성장에도 이익이 줄어듭니다.", affected_value_chain: "Production / Operation, Channel / Platform", evidence_to_watch: "ASP, 계약 구조, 가격 인상 발표, 고객 이탈률" },
      { variable: "정책/규제", current_signal: "정책은 수요를 앞당기거나 진입장벽을 높일 수 있지만, 동시에 비용과 불확실성도 만듭니다.", impact_on_industry: "보조금과 규제는 특정 기술·제품군의 성장 속도를 바꾸고, 기존 사업자의 투자 우선순위를 재배치합니다.", affected_value_chain: "전 밸류체인", evidence_to_watch: "보조금, 관세, 인허가, 환경·안전 규제 변경" }
    ],
    deep_dive: {
      headline_thesis: `${displayIndustry} 산업의 심층 분석은 수요 성장, 공급 제약, 가격 전가력, 정책 변화가 어느 방향으로 결합되는지에 달려 있습니다. 이 블록은 증권사 산업 리포트처럼 주요 이슈 플로우와 세부 섹터별 판단을 분리해 보여주기 위한 구조입니다.`,
      issue_flow: [
        { issue: "수요 사이클 변화", evidence: "최종 수요와 고객사 투자 집행이 산업 외형을 결정합니다.", industry_impact: "수요 확대가 가격 인상으로 이어지면 산업 이익률이 개선되지만, 공급 증가가 더 빠르면 매출 성장에도 마진은 제한됩니다." },
        { issue: "공급망 병목", evidence: "핵심 부품, 인력, 설비, 인증에서 병목이 발생할 수 있습니다.", industry_impact: "병목 구간은 단기적으로 협상력을 갖고 산업 내 profit pool을 흡수합니다." },
        { issue: "정책/규제 변화", evidence: "보조금, 관세, 인허가, 환경 규제가 투자 방향을 바꿉니다.", industry_impact: "정책은 신규 수요를 만들 수 있지만, 동시에 비용 부담과 경쟁 구도를 재편합니다." }
      ],
      subsector_breakdown: [
        { subsector: "상위 공급망", current_condition: "핵심 투입 요소의 가격과 공급 안정성이 산업 전체 비용 구조를 좌우합니다.", revenue_logic: "장기계약과 고부가 부품 판매에서 매출이 발생합니다.", cost_or_bottleneck: "원재료 가격, 설비투자, 인증 장벽이 주요 병목입니다.", outlook: "대체 공급이 어려운 품목은 수요 둔화기에도 상대적으로 마진을 방어할 수 있습니다." },
        { subsector: "제조/운영", current_condition: "가동률과 생산성이 이익률을 좌우하는 구간입니다.", revenue_logic: "판매량, ASP, 제품 믹스가 매출을 결정합니다.", cost_or_bottleneck: "고정비, 인건비, 에너지 비용이 손익 변동성을 만듭니다.", outlook: "수요 회복기에는 영업 레버리지가 크지만 둔화기에는 손익 훼손이 빠릅니다." },
        { subsector: "채널/서비스", current_condition: "고객 접점과 반복 매출을 가진 사업자가 실적 가시성을 확보합니다.", revenue_logic: "유통 마진, 플랫폼 수수료, 유지보수, 구독형 매출이 발생합니다.", cost_or_bottleneck: "고객 획득 비용과 물류비가 주요 부담입니다.", outlook: "고객 락인이 강하면 제조 단계보다 안정적인 이익을 만들 수 있습니다." }
      ],
      indicator_watch: [
        { indicator: "시장 규모/출하량", current_read: "산업 외형 성장 여부를 보여주는 1차 지표입니다.", why_it_moves_the_sector: "수요가 강하면 가동률과 투자 집행이 개선됩니다." },
        { indicator: "ASP/가격", current_read: "가격 전가력과 경쟁 강도를 보여줍니다.", why_it_moves_the_sector: "ASP가 유지되면 원가 상승기에도 마진 방어가 가능합니다." },
        { indicator: "가동률/수주잔고", current_read: "공급망의 실제 체감 수요를 보여줍니다.", why_it_moves_the_sector: "가동률 상승은 고정비 흡수와 이익률 개선으로 이어집니다." },
        { indicator: "정책/보조금/관세", current_read: "수요와 비용을 동시에 움직이는 외생 변수입니다.", why_it_moves_the_sector: "정책 변화는 특정 하위 시장의 성장 속도와 경쟁 구도를 바꿉니다." }
      ],
      scenario_analysis: [
        { case_name: "Bull Case", conditions: "수요 성장, 가격 전가, 공급 병목이 동시에 유지되는 경우", industry_result: "산업 외형과 마진이 함께 개선되고 병목 구간의 기업 가치가 재평가됩니다." },
        { case_name: "Base Case", conditions: "수요는 성장하지만 비용 부담과 경쟁이 일부 상쇄되는 경우", industry_result: "산업은 성장하지만 하위 섹터별 이익 차별화가 커집니다." },
        { case_name: "Bear Case", conditions: "수요 둔화와 공급 과잉, 정책 축소가 겹치는 경우", industry_result: "가격 하락과 가동률 둔화가 동시에 나타나 산업 마진이 훼손됩니다." }
      ],
      risk_factors: [
        { risk: "공급 과잉", transmission_path: "증설이 수요보다 빠르면 가격 하락과 재고 증가가 발생합니다.", early_signal: "가동률 하락, 재고 증가, ASP 하락" },
        { risk: "원가 상승", transmission_path: "원재료·인건비·물류비 상승이 비용 구조를 압박합니다.", early_signal: "원재료 가격 상승, 운임 상승, 가격 인상 실패" },
        { risk: "정책 변화", transmission_path: "보조금 축소나 규제 강화가 수요와 비용을 동시에 악화시킵니다.", early_signal: "정책 예산 축소, 인허가 지연, 관세 변경" }
      ]
    },
    report_implications: {
      core_conclusion: `${displayIndustry} 산업은 시장 규모 확대만으로 판단하기보다, 수요 성장의 이익이 어느 가치사슬 구간에 남는지를 중심으로 해석해야 합니다.`,
      implications: [
        "산업 외형이 커져도 공급 과잉이 빠르게 발생하면 가격 하락과 마진 훼손이 동시에 나타날 수 있습니다.",
        "핵심 부품, 기술 인증, 고객 락인, 유통 접점처럼 대체가 어려운 구간은 산업 내 profit pool을 흡수할 가능성이 높습니다.",
        "비용 구조가 무거운 산업에서는 매출 증가보다 고정비 흡수, 가동률, 원가 전가력이 실적 방향성을 더 잘 설명합니다."
      ],
      discussion_points: [
        "현재 산업 내 profit pool은 최종 제품, 핵심 부품, 채널, 서비스 중 어디에 집중되어 있는가?",
        "최근 수요 확대가 가격 상승으로 이어지는가, 아니면 공급 증가로 흡수되는가?",
        "정책·규제 변화가 수요 확대 요인인지 비용 증가 요인인지 구분할 수 있는가?"
      ]
    },
    sources: [
      { title: "산업 구조 예비 분석", publisher: "BIT Analysis", url: "", used_for: "산업 구조와 밸류체인 기본 프레임" },
      { title: industryContext.keywords.slice(0, 5).join(" / "), publisher: "검색 키워드", url: "", used_for: "상위 산업과 인접 세그먼트 탐색 기준" }
    ]
  };
}

function legacyIndustryChart(analysis, fallback) {
  return {
    marketSeries: normalizeSeries(analysis.chart && analysis.chart.marketSeries, fallback.chart.marketSeries),
    demandSeries: normalizeSeries(analysis.chart && analysis.chart.demandSeries, fallback.chart.demandSeries)
  };
}

function legacyIndustryFallback({ industry, scope, period, reportType }) {
  return {
    title: `${industry} 산업분석`,
    meta: ["기본 분석 요약", scope, period, reportType],
    chart: {
      marketSeries: [100, 112, 126, 143, 161],
      demandSeries: [100, 116, 134, 157, 184]
    }
  };
}

async function generateCompanyAnalysis({ financials, reportType }) {
  const payload = {
    model: openaiModel,
    instructions: [
      "You are a Korean corporate study analyst building a one-company study sheet for BIT Analysis.",
      "Write company analysis outputs, not instructions, methodology, interview answers, or resume sentences.",
      "The output should feel like an analyst-built company study sheet: basic company identity first, business model second, DART financials third.",
      "Do not expose reasoning steps, analysis process, checklist language, or phrases like '확인해야 합니다'. The user only wants finished insights.",
      "Start with what the company does, what it mainly sells, who it serves, what value or strategy it claims, how it makes money, and only then connect the numbers.",
      "Depth requirement: companyProfile fields must be specific to the company, not generic. Mention concrete product groups, customer/channel structure, strategy themes, and the business model in full sentences.",
      "Depth requirement: for Hyundai Motor, cover finished vehicles, Genesis/SUV mix, HEV/EV transition, global production/sales channels, captive finance, incentives, FX, and Hyundai Motor Group ecosystem when relevant.",
      "The user hates generic 'how to analyze' language. Give concrete business structure and concrete 핵심 해석.",
      "Use the DART financial packet and market data supplied by the app. Use your own company knowledge for business structure, but do not invent exact segment revenue percentages unless provided.",
      "사업적 강점 means reasons the company is competitively or structurally interesting for study purposes. It is not investment advice.",
      "For graphInsights, interpret the supplied five-year DART numbers: observed change, meaning, and 시사점.",
      "If a fact is uncertain, write it directionally and make the uncertainty explicit in sourceNote, not in every sentence.",
      "Return only valid JSON matching the requested schema."
    ].join("\n"),
    input: JSON.stringify({
      report_type: reportType,
      company: financials.company,
      request: financials.request,
      periods: compactFinancialPeriods(financials.periods),
      quality: financials.quality,
      market_data: financials.market_data,
      valuation: financials.valuation,
      cash_flow_pattern: financials.cash_flow_pattern,
      existing_numeric_insights: financials.insights,
      business_context: companyStudyContext(financials.company && (financials.company.name || financials.company.corp_name))
    }),
    max_output_tokens: 6000,
    text: {
      format: {
        type: "json_schema",
        name: "company_analysis",
        strict: true,
        schema: companyAnalysisSchema()
      }
    }
  };

  const response = await fetchWithTimeout(OPENAI_RESPONSES_URL, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${openaiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  }, 60000);

  const data = await response.json().catch(() => null);
  if (!response.ok) {
    const message = data && data.error && data.error.message ? data.error.message : `HTTP ${response.status}`;
    throw new Error(message);
  }

  const text = extractResponseText(data);
  if (!text) throw new Error("OpenAI response did not contain output text.");
  const parsed = JSON.parse(text);
  return normalizeCompanyAnalysis(parsed, financials, reportType);
}

function companyStudyContext(name) {
  const normalized = normalizeQuery(name || "");
  if (normalized.includes("현대자동차") || normalized.includes("현대차")) {
    return {
      identity: "현대자동차는 완성차 제조·판매를 중심으로 금융, 부품/서비스, 미래 모빌리티 투자를 함께 운영하는 글로벌 OEM이다.",
      core_business: [
        "승용차/SUV/상용차 제조와 글로벌 판매",
        "Genesis 등 고부가 브랜드와 SUV 중심 믹스 개선",
        "하이브리드, 전기차, 수소전기차 등 전동화 라인업",
        "현대캐피탈 등 자동차 금융과 리스/할부 기반 판매 지원",
        "부품, AS, 커넥티드카, 소프트웨어 기반 서비스"
      ],
      revenue_logic: "매출은 판매대수, ASP, 차급/브랜드 믹스, 지역 믹스, 금융/리스 수익에서 나온다. 이익은 인센티브, 환율, 원재료, 물류비, 가동률, 금융부문 충당/조달비용에 민감하다.",
      current_focus: "최근 분석 초점은 단순 판매대수보다 SUV/Genesis/HEV 비중, 미국·인도 등 지역 믹스, 전기차 수요 둔화 대응, 금융부문 건전성, 대규모 투자에도 현금흐름을 유지하는 능력이다.",
      risk_context: "자동차 금융 부채와 제조업 차입금을 같은 방식으로 해석하면 안정성 판단이 왜곡된다. 금융부문은 판매를 돕지만 금리, 잔존가치, 연체율 리스크를 동반한다."
    };
  }
  if (normalized.includes("삼성전자")) {
    return {
      identity: "삼성전자는 반도체, 스마트폰, 디스플레이, 가전으로 구성된 글로벌 종합 전자기업이다.",
      core_business: ["메모리 반도체", "파운드리/시스템 반도체", "MX 스마트폰", "VD/가전", "디스플레이"],
      revenue_logic: "매출은 메모리 가격/출하량, 스마트폰 판매량/ASP, 프리미엄 가전 믹스에서 나오며 이익은 반도체 업황과 고부가 제품 믹스에 민감하다.",
      current_focus: "HBM, AI 서버향 메모리, 파운드리 수율, 스마트폰 프리미엄 믹스, 대규모 CAPEX 회수 가능성이 핵심이다.",
      risk_context: "메모리 사이클, 파운드리 경쟁, CAPEX 부담, 환율과 세트 수요 둔화가 주요 리스크다."
    };
  }
  return {
    identity: "사업보고서의 사업부문 설명과 DART 재무를 함께 연결해 읽는 기업이다.",
    core_business: [],
    revenue_logic: "매출은 주력 제품/서비스의 가격, 물량, 고객 믹스로 형성되고 이익은 비용 구조와 가격 전가력에 의해 결정된다.",
    current_focus: "기업 기본정보, 사업 구조, 매출/이익/현금흐름의 연결성이 핵심이다.",
    risk_context: "사업부문 원문과 재무제표 계정의 연결이 부족하면 해석이 얕아진다."
  };
}

function companyAnalysisSchema() {
  return {
    type: "object",
    additionalProperties: false,
    required: [
      "title",
      "meta",
      "companyProfile",
      "companyOverview",
      "soWhatLead",
      "businessStructure",
      "revenueCost",
      "financialSoWhat",
      "investmentHighlights",
      "riskFactors",
      "graphInsights",
      "onePageBrief",
      "sourceNote"
    ],
    properties: {
      title: { type: "string" },
      meta: { type: "array", minItems: 3, maxItems: 6, items: { type: "string" } },
      companyProfile: {
        type: "object",
        additionalProperties: false,
        required: ["identity", "coreBusiness", "mainProducts", "customersAndChannels", "statedValue", "businessModel", "currentFocus"],
        properties: {
          identity: { type: "string" },
          coreBusiness: { type: "string" },
          mainProducts: { type: "string" },
          customersAndChannels: { type: "string" },
          statedValue: { type: "string" },
          businessModel: { type: "string" },
          currentFocus: { type: "string" }
        }
      },
      companyOverview: { type: "string" },
      soWhatLead: { type: "string" },
      businessStructure: {
        type: "array",
        minItems: 4,
        maxItems: 5,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["segment", "whatItDoes", "revenueLogic", "marginLogic", "soWhat"],
          properties: {
            segment: { type: "string" },
            whatItDoes: { type: "string" },
            revenueLogic: { type: "string" },
            marginLogic: { type: "string" },
            soWhat: { type: "string" }
          }
        }
      },
      revenueCost: {
        type: "array",
        minItems: 3,
        maxItems: 5,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["item", "mechanism", "soWhat"],
          properties: {
            item: { type: "string" },
            mechanism: { type: "string" },
            soWhat: { type: "string" }
          }
        }
      },
      financialSoWhat: {
        type: "array",
        minItems: 5,
        maxItems: 8,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["metric", "value", "interpretation", "soWhat"],
          properties: {
            metric: { type: "string" },
            value: { type: "string" },
            interpretation: { type: "string" },
            soWhat: { type: "string" }
          }
        }
      },
      investmentHighlights: { type: "array", minItems: 3, maxItems: 5, items: { type: "string" } },
      riskFactors: { type: "array", minItems: 3, maxItems: 5, items: { type: "string" } },
      graphInsights: {
        type: "array",
        minItems: 4,
        maxItems: 5,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["title", "observation", "interpretation", "soWhat"],
          properties: {
            title: { type: "string" },
            observation: { type: "string" },
            interpretation: { type: "string" },
            soWhat: { type: "string" }
          }
        }
      },
      onePageBrief: { type: "array", minItems: 4, maxItems: 7, items: { type: "string" } },
      sourceNote: { type: "string" }
    }
  };
}

function compactFinancialPeriods(periods) {
  return periods.map(period => ({
    year: period.year,
    fs_div: period.fs_div,
    metrics: period.metrics
  }));
}

function normalizeCompanyAnalysis(analysis, financials, reportType) {
  const fallback = buildCompanyAnalysisFallback(financials, reportType);
  return {
    ...fallback,
    ...analysis,
    title: analysis.title || fallback.title,
    meta: Array.isArray(analysis.meta) && analysis.meta.length ? analysis.meta : fallback.meta,
    companyProfile: analysis.companyProfile || fallback.companyProfile,
    businessStructure: normalizeArray(analysis.businessStructure, fallback.businessStructure),
    revenueCost: normalizeArray(analysis.revenueCost, fallback.revenueCost),
    financialSoWhat: normalizeArray(analysis.financialSoWhat, fallback.financialSoWhat),
    investmentHighlights: normalizeArray(analysis.investmentHighlights, fallback.investmentHighlights),
    riskFactors: normalizeArray(analysis.riskFactors, fallback.riskFactors),
    graphInsights: normalizeArray(analysis.graphInsights, fallback.graphInsights),
    onePageBrief: normalizeArray(analysis.onePageBrief, fallback.onePageBrief)
  };
}

function normalizeArray(value, fallback) {
  return Array.isArray(value) && value.length ? value : fallback;
}

function buildCompanyAnalysisFallback(financials, reportType) {
  const company = financials.company || {};
  const name = company.name || company.corp_name || company.corp_code || "해당 기업";
  const latest = [...(financials.periods || [])].reverse().find(period => period.metrics && Object.keys(period.metrics).length) || { year: financials.request && financials.request.base_year, metrics: {} };
  const first = (financials.periods || []).find(period => period.metrics && Object.keys(period.metrics).length) || latest;
  const m = latest.metrics || {};
  const firstMetrics = first.metrics || {};
  const revenueGrowth = Number.isFinite(m.revenue) && Number.isFinite(firstMetrics.revenue) && firstMetrics.revenue !== 0
    ? (m.revenue / firstMetrics.revenue) - 1
    : null;
  const opMargin = Number.isFinite(m.revenue) && Number.isFinite(m.operating_income) && m.revenue !== 0 ? m.operating_income / m.revenue : null;
  const cfoToOp = Number.isFinite(m.operating_cash_flow) && Number.isFinite(m.operating_income) && m.operating_income !== 0 ? m.operating_cash_flow / m.operating_income : null;
  const pattern = financials.cash_flow_pattern || classifyCashFlowPattern(m);
  const valuation = financials.valuation || {};

  return {
    title: `${name} 기업분석`,
    meta: ["기업분석", reportType, `${latest.year || ""} DART`, financials.quality && financials.quality.score ? `${financials.quality.score} quality` : "DART"],
    companyProfile: {
      identity: `${name}은 상장 공시와 사업보고서를 기준으로 사업 구조와 재무 성과를 함께 읽어야 하는 분석 대상 기업입니다.`,
      coreBusiness: "주요 사업은 회사가 반복적으로 매출을 만드는 제품/서비스군, 생산·운영 역량, 고객 기반으로 구성됩니다.",
      mainProducts: "주요 제품과 서비스는 사업보고서의 사업부문 설명을 기준으로 보강됩니다.",
      customersAndChannels: "고객과 채널은 B2B/B2C, 국내/해외, 직접 판매/유통 구조로 구분해 읽습니다.",
      statedValue: "표방 가치는 회사가 IR, 사업보고서, 홈페이지에서 반복적으로 강조하는 기술, 품질, 고객, 지속가능성, 혁신 방향으로 정리됩니다.",
      businessModel: `돈을 버는 방식은 매출 ${formatWon(m.revenue)}를 만드는 판매량·가격·제품 믹스와 EBIT margin ${formatPercent(opMargin)}를 만드는 비용 구조의 조합입니다.`,
      currentFocus: `최근 초점은 외형보다 영업현금흐름 ${formatWon(m.operating_cash_flow)}와 순차입금 ${formatWon(m.net_debt)}가 보여주는 사업 체력입니다.`
    },
    companyOverview: `${name}은 먼저 '무엇을 팔고 누구에게 가치를 제공하는가'를 이해한 뒤, DART 숫자로 그 사업이 실제로 얼마나 돈을 벌고 있는지 연결해서 읽는 기업입니다. 최신 DART 기준 매출은 ${formatWon(m.revenue)}, EBIT는 ${formatWon(m.operating_income)}입니다.`,
    soWhatLead: `${name}을 볼 때는 기업 정체성, 주력 제품/서비스, 고객 기반, 돈 버는 방식이 먼저이고, 재무는 그 사업 설명이 실제 성과로 이어졌는지 검증하는 근거입니다.`,
    businessStructure: [
      { segment: "핵심 제품/서비스", whatItDoes: "회사의 주력 매출을 만드는 제품과 서비스군", revenueLogic: "가격, 물량, 고객 믹스가 매출을 결정", marginLogic: "고정비 흡수와 제품 믹스가 EBIT margin을 결정", soWhat: "매출 성장만으로는 부족하고 고마진 제품 비중이 함께 올라가야 재평가됩니다." },
      { segment: "고객/채널", whatItDoes: "B2B/B2C 고객과 판매 채널", revenueLogic: "반복 고객과 장기 계약은 실적 가시성을 높임", marginLogic: "고객 집중도와 가격 전가력이 마진 방어력을 결정", soWhat: "고객 락인이 강할수록 업황 둔화에도 매출과 현금흐름이 덜 흔들립니다." },
      { segment: "생산/운영", whatItDoes: "제조, 서비스 운영, CAPEX 기반", revenueLogic: "가동률 상승은 매출 확대와 연결", marginLogic: "감가상각과 인건비 등 고정비 부담이 큼", soWhat: "CAPEX가 큰 기업은 호황기에 이익 레버리지가 크지만 불황기에는 손익 훼손이 빠릅니다." },
      { segment: "재무 체력", whatItDoes: "현금, 차입금, 영업현금흐름", revenueLogic: "재무 여력은 성장투자와 주주환원의 원천", marginLogic: "이자비용과 투자 부담이 순이익을 좌우", soWhat: `순차입금 ${formatWon(m.net_debt)} 구조는 다음 투자 사이클을 버틸 수 있는지의 1차 판단 기준입니다.` }
    ],
    revenueCost: [
      { item: "Revenue", mechanism: `최근 매출은 ${formatWon(m.revenue)}이며 5개년 누적 변화는 ${formatPercent(revenueGrowth)}입니다.`, soWhat: "외형 성장은 확인되더라도 EBIT margin이 같이 개선되어야 질 좋은 성장입니다." },
      { item: "Cost", mechanism: "원재료, 인건비, 감가상각, R&D, 판관비가 EBIT를 압박합니다.", soWhat: `현재 EBIT margin ${formatPercent(opMargin)}은 비용 전가력과 가동률의 결과입니다.` },
      { item: "Cash Conversion", mechanism: `영업현금흐름/EBIT는 ${formatMultiple(cfoToOp)}입니다.`, soWhat: "회계상 이익보다 현금 전환이 강하면 CAPEX와 차입 상환을 내부 현금으로 감당할 여지가 큽니다." },
      { item: "Valuation", mechanism: `PER ${formatMultiple(valuation.per)}, EV/EBITDA ${formatMultiple(valuation.ev_ebitda)}입니다.`, soWhat: "밸류에이션은 이익 회복이 이미 주가에 얼마나 반영됐는지 보는 체크포인트입니다." }
    ],
    financialSoWhat: [
      { metric: "매출액", value: formatWon(m.revenue), interpretation: `5개년 누적 변화 ${formatPercent(revenueGrowth)}`, soWhat: "성장률 자체보다 EBIT margin이 동반 개선되는지가 중요합니다." },
      { metric: "EBIT", value: formatWon(m.operating_income), interpretation: `EBIT margin ${formatPercent(opMargin)}`, soWhat: "본업 경쟁력은 매출보다 EBIT margin에서 먼저 드러납니다." },
      { metric: "영업현금흐름", value: formatWon(m.operating_cash_flow), interpretation: `CFO/EBIT ${formatMultiple(cfoToOp)}`, soWhat: "현금전환이 강하면 투자와 차입 상환을 버틸 체력이 있습니다." },
      { metric: "순차입금", value: formatWon(m.net_debt), interpretation: Number.isFinite(m.net_debt) && m.net_debt < 0 ? "순현금" : "순차입", soWhat: "재무구조는 다운사이드 방어와 다음 성장투자의 여력을 결정합니다." },
      { metric: "현금흐름 패턴", value: pattern.type, interpretation: pattern.description, soWhat: "영업/투자/재무 CF 조합이 회사의 현재 사이클 위치를 보여줍니다." }
    ],
    investmentHighlights: [
      "주력 사업의 고객 기반과 제품 믹스가 명확할수록 매출의 반복성과 가격 전가력이 높아집니다.",
      `최신 매출 ${formatWon(m.revenue)}와 EBIT ${formatWon(m.operating_income)}는 사업모델이 실제 손익으로 연결되는 정도를 보여줍니다.`,
      `영업현금흐름 ${formatWon(m.operating_cash_flow)}는 회사가 벌어들인 이익을 실제 현금으로 전환하는 힘을 보여줍니다.`
    ],
    riskFactors: [
      "매출 성장과 이익률 개선이 분리될 경우 외형 성장에도 밸류에이션 확장이 제한됩니다.",
      "CAPEX 또는 운전자본 부담이 커지면 영업현금흐름이 약화될 수 있습니다.",
      "시장 기대가 이미 PER에 반영되어 있다면 실적 개선에도 주가 업사이드가 제한될 수 있습니다."
    ],
    graphInsights: [
      {
        title: "매출액 변화 추이",
        observation: `최근 매출은 ${formatWon(m.revenue)}이고 5개년 누적 변화는 ${formatPercent(revenueGrowth)}입니다.`,
        interpretation: "외형 성장만으로는 충분하지 않고 EBIT와 EBITDA가 같은 방향으로 개선되는지가 중요합니다.",
        soWhat: "매출이 늘어도 이익률이 따라오지 않으면 성장의 질은 낮게 평가됩니다."
      },
      {
        title: "EBITDA 변화 추이",
        observation: `최근 EBIT는 ${formatWon(m.operating_income)}, EBITDA는 ${formatWon(m.ebitda)}입니다.`,
        interpretation: "EBITDA는 감가상각 부담을 제거한 현금성 영업 체력을 보여줍니다.",
        soWhat: "EBITDA 회복은 CAPEX가 큰 기업이 다음 투자 사이클을 버틸 수 있는지 판단하는 핵심 신호입니다."
      },
      {
        title: "영업현금흐름 변화 추이",
        observation: `최근 영업현금흐름은 ${formatWon(m.operating_cash_flow)}이고 CFO/EBIT는 ${formatMultiple(cfoToOp)}입니다.`,
        interpretation: "회계상 이익이 실제 현금으로 전환되는지가 이익의 질을 가릅니다.",
        soWhat: "영업현금흐름이 강하면 투자, 차입 상환, 주주환원을 내부 현금으로 감당할 가능성이 높습니다."
      },
      {
        title: "순차입금 변화 추이",
        observation: `최근 순차입금은 ${formatWon(m.net_debt)}입니다.`,
        interpretation: "순차입금은 현금 여력과 차입 부담을 동시에 보여주는 안정성 지표입니다.",
        soWhat: "순현금 또는 낮은 순차입 구조는 업황 둔화와 투자 확대를 버티는 방어력입니다."
      }
    ],
    onePageBrief: [
      `${name}: ${latest.year || "최신"}년 매출 ${formatWon(m.revenue)}, EBIT ${formatWon(m.operating_income)}.`,
      `핵심 해석: EBIT margin ${formatPercent(opMargin)}과 CFO/EBIT ${formatMultiple(cfoToOp)}가 성장의 질을 판단하는 핵심입니다.`,
      `재무 체력: 순차입금 ${formatWon(m.net_debt)}, 부채비율 ${formatPercent(m.debt_ratio)}.`,
      `현금흐름: ${pattern.type}. ${pattern.description}`,
      `밸류에이션: PER ${formatMultiple(valuation.per)}, EV/EBITDA ${formatMultiple(valuation.ev_ebitda)}.`
    ],
    sourceNote: "DART 재무/시장 데이터 기반 요약입니다. 사업부별 제품/고객 세부 정보는 사업보고서 원문 리서치로 추가 보강됩니다."
  };
}

function buildLegacyIndustryFallbackFull({ industry, scope, period, reportType }) {
  const isAiSemi = /AI|반도체|HBM|semiconductor/i.test(industry);
  const isEnergy = /정유|화학|에너지|유틸|전력|LNG|가스/i.test(industry);
  const isAuto = /자동차|모빌리티|전기차|하이브리드|HEV/i.test(industry);
  const isTelecom = /통신|5G|6G|네트워크|장비|AI\s*RAN/i.test(industry);

  const preset = isAiSemi ? {
    oneLine: `${industry} 산업은 최근 AI 서버 투자가 GPU 중심에서 HBM, 첨단패키징, 선단 파운드리 병목으로 확산되며 공급 제약 구간에 이익이 집중되는 상황입니다. ${scope} 범위에서는 CSP의 CAPEX 지속성, HBM 고객 인증, 패키징 캐파가 수익성을 가릅니다.`,
    stance: "긍정. 수요는 AI 학습에서 추론·Agentic AI로 넓어지고 있고, 공급은 HBM/패키징 병목 때문에 단기간에 쉽게 풀리지 않는 구조입니다.",
    thesis: [
      "AI 반도체의 이익 풀은 GPU 설계사에만 머물지 않고 HBM, 패키징, 테스트, 전력/기판 등 병목 공급망으로 이동하고 있습니다.",
      "HBM은 범용 DRAM보다 고객 인증과 품질 안정성이 중요해 가격 하락 압력이 상대적으로 낮고, 선두 업체의 마진 방어력이 강합니다.",
      "후발 업체는 수요 성장보다 수율·인증·패키징 확보 속도가 느리면 업황 호조를 실적으로 온전히 흡수하기 어렵습니다."
    ],
    whyNow: "AI 모델 경쟁이 계속되면서 클라우드 사업자의 투자 집행이 반도체 수요를 지지하고 있습니다. 동시에 전력, 패키징, HBM 공급이 병목으로 남아 있어 단순 판매량보다 공급 우위가 프리미엄을 만드는 구간입니다.",
    definition: "AI 반도체는 대규모 AI 연산을 처리하는 GPU, ASIC, HBM, 인터커넥트, 첨단패키징, 파운드리 생태계를 포함합니다.",
    variables: ["CSP AI CAPEX", "HBM 공급과 고객 인증", "첨단패키징 캐파", "파운드리 수율", "서버 전력/냉각 인프라"],
    recentChange: "최근 시장의 관심은 AI 수요 자체보다 병목 부품의 공급 여력과 수익성으로 이동했습니다. GPU 출하가 늘어도 HBM·패키징이 따라오지 못하면 최종 서버 공급이 제한됩니다.",
    valueChain: [
      { stage: "AI Accelerator", detail: "GPU와 ASIC 설계, 플랫폼 생태계, 소프트웨어 락인", implication: "NVIDIA 같은 선도 설계사는 생태계 지배력으로 높은 ASP와 마진을 유지합니다." },
      { stage: "HBM/Memory", detail: "HBM3E·차세대 HBM, DRAM 선단 공정, 고객 인증", implication: "SK하이닉스와 삼성전자처럼 HBM 품질과 공급 능력을 가진 업체가 메모리 사이클보다 높은 프리미엄을 받습니다." },
      { stage: "Foundry/Packaging", detail: "선단 파운드리, CoWoS류 첨단패키징, 기판·테스트", implication: "TSMC와 패키징 밸류체인은 AI 서버 출하의 병목이기 때문에 가격 협상력이 강합니다." },
      { stage: "Cloud/Data Center", detail: "CSP, AI Native cloud, Sovereign AI, 데이터센터 전력", implication: "최종 수요는 클라우드 CAPEX와 전력 확보에 의해 결정되며, 투자 지연은 반도체 주문 조정으로 이어집니다." }
    ],
    catalysts: [
      { event: "글로벌 AI CAPEX 확대", impact: "GPU와 HBM 주문 가시성을 높이고 장기 공급계약을 유도합니다.", soWhat: "HBM 선두 업체와 패키징 병목 업체의 실적 가시성이 설계사 못지않게 부각됩니다." },
      { event: "차세대 GPU/가속기 전환", impact: "메모리 용량과 대역폭 요구가 커지며 HBM 탑재량이 증가합니다.", soWhat: "제품 세대 전환은 단순 교체 수요가 아니라 HBM ASP와 믹스를 끌어올리는 이벤트입니다." },
      { event: "전력·냉각 제약", impact: "데이터센터 증설 속도를 제한하고 고효율 칩 수요를 자극합니다.", soWhat: "AI 반도체 경쟁력은 연산 성능뿐 아니라 전력 효율과 시스템 공급 능력으로 확장됩니다." }
    ],
    revenueCost: [
      { item: "Revenue", mechanism: "AI 서버 투자, GPU/ASIC 출하, HBM 탑재량, 장기 공급계약이 매출을 만듭니다.", soWhat: "HBM 비중이 높아지는 기업은 같은 bit growth에서도 매출과 이익의 질이 달라집니다." },
      { item: "Cost", mechanism: "선단 공정 CAPEX, 수율 안정화, 패키징 투자, 테스트 비용이 비용 부담입니다.", soWhat: "수율이 안정된 선두 업체는 고정비를 흡수하지만 후발 업체는 호황에도 마진 회복이 늦습니다." },
      { item: "Margin Driver", mechanism: "고객 인증, 공급 부족, 제품 세대 전환, 패키징 캐파가 마진을 좌우합니다.", soWhat: "AI 반도체의 프리미엄은 시장 성장률보다 병목 구간의 희소성에서 발생합니다." }
    ],
    keyPlayers: [
      { name: "NVIDIA", position: "GPU와 CUDA 생태계 중심의 플랫폼 리더", implication: "AI 투자 사이클의 방향성을 결정하지만 밸류에이션 부담도 가장 먼저 반영됩니다." },
      { name: "SK하이닉스", position: "HBM 선두 공급자", implication: "HBM 프리미엄이 DRAM 사이클 민감도를 완충하며 메모리 업체 재평가의 핵심 축입니다." },
      { name: "삼성전자", position: "메모리·파운드리·패키징을 동시에 보유한 종합 공급자", implication: "HBM 인증과 파운드리 수율 회복이 확인될수록 업사이드가 커지는 구조입니다." },
      { name: "TSMC", position: "선단 파운드리와 첨단패키징 병목의 핵심", implication: "AI 칩 출하 확대의 병목을 쥐고 있어 가격 협상력과 투자 가시성이 높습니다." }
    ],
    talkingPoints: [
      "AI 반도체는 수요 고성장 산업이 아니라 병목 공급망 프리미엄 산업으로 바뀌고 있습니다.",
      "HBM과 첨단패키징은 AI 서버 출하를 제한하는 구간이기 때문에 최종 칩보다 높은 협상력을 갖는 순간이 생깁니다.",
      "한국 반도체의 핵심 해석은 메모리 사이클 반등이 아니라 HBM을 통해 AI 인프라 밸류체인에 직접 편입된다는 점입니다."
    ],
    questions: ["HBM 공급 부족이 언제부터 가격 하락 압력으로 바뀔 수 있는가?", "AI CAPEX가 학습에서 추론으로 이동할 때 필요한 칩과 메모리 구조는 어떻게 달라지는가?"],
    chart: { marketSeries: [100, 132, 171, 224, 292], demandSeries: [100, 138, 196, 278, 390] }
  } : isEnergy ? {
    oneLine: `${industry} 산업은 원유·LNG 가격 변동과 에너지 안보 이슈가 맞물리며 정제마진, 발전원가, 석화 스프레드가 동시에 흔들리는 상황입니다. ${scope} 범위에서는 원료 조달 안정성과 비용 전가력이 기업별 실적을 가릅니다.`,
    stance: "중립. 에너지 가격 상승은 정유와 자원 보유 기업에는 우호적이지만, 나프타·LNG를 원료로 쓰는 석화와 발전 사업자에는 마진 압박으로 작동합니다.",
    thesis: [
      "정유는 공급 차질과 제품 부족이 정제마진을 지지하지만, 석유화학은 나프타 원가 상승과 중국 증설 부담이 겹칩니다.",
      "LNG와 전력 가격 변동성은 유틸리티의 원가 부담과 정책 리스크를 동시에 키웁니다.",
      "재생에너지 정책은 성장 모멘텀이지만 경쟁입찰과 상한가격은 사업자 마진을 제한합니다."
    ],
    whyNow: "에너지 가격과 정책이 동시에 움직이면 같은 에너지 밸류체인 안에서도 수혜와 피해가 갈립니다. 에너지 독립성, 원료 소싱, 장기계약 비중이 주가 차별화의 핵심입니다.",
    definition: "에너지/정유화학 산업은 원유·가스 조달, 정제, 석유화학 원료, 전력·열 공급, 재생에너지 개발을 포괄합니다.",
    variables: ["두바이/WTI 유가", "나프타·LNG 가격", "정제마진", "전력 도매가격", "재생에너지 입찰 단가"],
    recentChange: "최근 업황은 고유가 자체보다 공급 차질이 어떤 제품의 가격 전가를 가능하게 하는지에 초점이 맞춰져 있습니다.",
    valueChain: [
      { stage: "Upstream", detail: "원유, 천연가스, LPG, LNG 장기계약", implication: "자원 접근성이 있는 기업은 가격 급등기에 조달 안정성과 이익 방어력을 얻습니다." },
      { stage: "Refining", detail: "정제설비, 항공유·경유·휘발유 생산", implication: "제품 공급 부족은 정제마진을 끌어올려 정유사의 이익 레버리지로 이어집니다." },
      { stage: "Petrochemical", detail: "나프타, 에틸렌, 합성수지, 다운스트림 소재", implication: "원료 가격 상승을 제품가에 전가하지 못하면 스프레드가 축소됩니다." },
      { stage: "Power/Renewable", detail: "발전, ESS, 태양광·풍력, 계통 접속", implication: "정책 지원은 성장성을 만들지만 입찰 경쟁은 프로젝트 마진을 낮춥니다." }
    ],
    catalysts: [
      { event: "원유/LNG 공급 차질", impact: "정제마진과 발전원가가 동시에 상승합니다.", soWhat: "정유·자원 보유 기업은 수혜, 원료 구매 비중이 큰 석화·유틸리티는 비용 부담이 커집니다." },
      { event: "재생에너지 보급 정책", impact: "태양광·풍력·ESS 투자를 확대합니다.", soWhat: "기자재와 계통 인프라는 수혜지만 개발사업자는 경쟁입찰로 마진이 제한됩니다." },
      { event: "석화 구조조정", impact: "과잉설비 축소가 스프레드 회복의 조건이 됩니다.", soWhat: "구조조정 속도가 중국 증설 부담보다 빨라야 본격적인 업황 회복이 가능합니다." }
    ],
    revenueCost: [
      { item: "Revenue", mechanism: "정제제품 판매, 전력/열 판매, 석화제품 스프레드, 재생 프로젝트 매출이 외형을 구성합니다.", soWhat: "같은 유가 상승도 정유에는 매출과 마진 개선, 석화에는 원가 부담으로 작동합니다." },
      { item: "Cost", mechanism: "원유, 나프타, LNG, 탄소비용, 금융비용이 핵심 비용입니다.", soWhat: "조달 구조가 약한 기업은 가격 상승기에 매출이 늘어도 이익이 훼손됩니다." },
      { item: "Margin Driver", mechanism: "정제마진, 스프레드, SMP, 장기계약, 입찰 단가가 마진을 결정합니다.", soWhat: "가격 전가 가능한 구간과 정부가 단가를 누르는 구간을 분리해야 합니다." }
    ],
    keyPlayers: [
      { name: "SK이노베이션/S-Oil", position: "정유와 에너지 공급 안정성의 핵심 플레이어", implication: "제품 부족과 정제마진 상승기에 이익 레버리지가 큽니다." },
      { name: "롯데케미칼/LG화학", position: "석유화학 다운스트림 주요 기업", implication: "구조조정과 고부가 제품 전환 없이는 원가 상승 압력이 부담입니다." },
      { name: "한국전력/발전·가스 기업", position: "전력·가스 조달과 정책 요금의 접점", implication: "원가 상승과 요금 규제의 간극이 실적 변동성을 만듭니다." }
    ],
    talkingPoints: [
      "에너지 산업은 고유가가 모두에게 좋은 산업이 아니라, 원료를 파는 쪽과 사는 쪽의 실적이 갈리는 산업입니다.",
      "정유는 제품 공급 부족이 정제마진으로 연결될 때 투자 매력이 커지고, 석화는 나프타 가격보다 스프레드 회복이 핵심입니다.",
      "재생에너지 정책은 성장성을 만들지만, 입찰 구조는 사업자 마진을 낮추는 양면성이 있습니다."
    ],
    questions: ["고유가가 정유와 석화에 서로 다르게 작동하는 이유는 무엇인가?", "재생에너지 보급 확대가 왜 모든 사업자에게 동일한 수혜가 아닌가?"],
    chart: { marketSeries: [100, 108, 122, 140, 158], demandSeries: [100, 111, 125, 137, 151] }
  } : isAuto ? {
    oneLine: `${industry} 산업은 고유가와 전동화 보조금 변화가 맞물리며 BEV 일변도에서 HEV 중심의 수요 재배분이 나타나는 상황입니다. ${scope} 범위에서는 하이브리드 공급 능력과 인센티브 통제가 마진을 가릅니다.`,
    stance: "긍정적 선별. 완성차 전체 수요는 둔화될 수 있지만 HEV와 고부가 차종을 즉시 공급할 수 있는 업체는 비용 부담을 가격과 믹스로 흡수할 수 있습니다.",
    thesis: [
      "BEV 재고와 인센티브 부담이 커지는 동안 HEV는 고유가와 보조금 축소의 대안 수요를 흡수합니다.",
      "현대차·기아처럼 SUV/HEV 라인업과 현지 생산을 동시에 보유한 업체는 관세·물류비 부담을 상대적으로 완화합니다.",
      "부품사는 완성차 생산 믹스가 HEV로 이동할 때 열관리, 구동계, 전장 부품에서 수혜가 생깁니다."
    ],
    whyNow: "자동차 업종의 핵심은 판매대수 증가보다 인센티브를 낮게 유지하면서 고마진 차종을 팔 수 있는지입니다. HEV 쇼티지는 하반기 수익성 방어 논리로 작동합니다.",
    definition: "자동차 산업은 완성차, 부품, 전장, 금융, 서비스가 결합된 경기민감 소비재 산업입니다.",
    variables: ["미국/유럽 신차 수요", "HEV/BEV 믹스", "인센티브", "원재료·물류비", "관세와 현지 생산"],
    recentChange: "최근 전동화 논리는 BEV 침투율 자체보다 수익성 있는 전동화, 특히 HEV 공급 능력으로 이동하고 있습니다.",
    valueChain: [
      { stage: "OEM", detail: "현대차, 기아, Toyota, Honda 등 완성차 생산·판매", implication: "HEV 라인업과 현지 생산 능력이 마진 방어력을 만듭니다." },
      { stage: "Powertrain/HEV", detail: "엔진, 모터, 배터리, 변속, 열관리", implication: "HEV 수요 확대는 기존 내연기관 부품과 전장 부품 모두에 수혜를 줍니다." },
      { stage: "BEV/Battery", detail: "전기차 플랫폼, 배터리, 충전 인프라", implication: "재고와 인센티브 부담이 커지면 단기 수익성은 HEV 대비 약해집니다." },
      { stage: "Dealer/Finance", detail: "판매망, 할부/리스, 중고차 잔존가치", implication: "잔존가치와 금융비용은 소비자 실구매가와 브랜드 수익성을 동시에 흔듭니다." }
    ],
    catalysts: [
      { event: "고유가 지속", impact: "연비 민감도가 높아지며 HEV 수요가 확대됩니다.", soWhat: "HEV 공급 가능한 완성차는 인센티브를 덜 쓰고 점유율을 확대할 수 있습니다." },
      { event: "BEV 보조금 축소", impact: "전기차 구매 부담이 높아지고 BEV 재고 부담이 커집니다.", soWhat: "BEV 중심 업체보다 HEV/ICE/BEV 믹스를 유연하게 조절하는 업체가 유리합니다." },
      { event: "관세·현지 생산 이슈", impact: "수입차 원가와 가격 정책에 영향을 줍니다.", soWhat: "미국 현지 생산 확대는 관세 부담과 고정비 부담을 동시에 낮춥니다." }
    ],
    revenueCost: [
      { item: "Revenue", mechanism: "판매대수, ASP, SUV/HEV 믹스, 금융 매출이 외형을 만듭니다.", soWhat: "판매대수가 정체돼도 HEV와 SUV 믹스가 개선되면 매출과 마진이 동시에 방어됩니다." },
      { item: "Cost", mechanism: "원재료, 물류, 배터리, 관세, 인센티브가 비용 부담입니다.", soWhat: "인센티브가 낮은 차종을 팔 수 있는 업체가 업황 둔화기에도 이익을 지킵니다." },
      { item: "Margin Driver", mechanism: "고부가 차종 믹스, 현지 생산, 플랫폼 공용화, 잔존가치가 마진을 좌우합니다.", soWhat: "자동차의 단기 핵심 해석은 전동화 속도보다 수익성 있는 전동화입니다." }
    ],
    keyPlayers: [
      { name: "현대차/기아", position: "SUV와 HEV 라인업을 확대하는 글로벌 완성차", implication: "미국 HEV 수요 확대를 흡수하면 마진 방어와 점유율 확대가 동시에 가능합니다." },
      { name: "Toyota/Honda", position: "HEV 선도 업체", implication: "HEV 수요 확대의 기준점이자 현대차그룹이 점유율을 가져와야 할 대상입니다." },
      { name: "현대모비스/한온시스템", position: "전동화·열관리·핵심 부품 공급자", implication: "차종 믹스 변화가 부품 수주와 마진 개선으로 연결될 수 있습니다." }
    ],
    talkingPoints: [
      "자동차 산업의 핵심은 BEV 전환 속도가 아니라 수익성 있는 전동화 믹스입니다.",
      "HEV 쇼티지가 발생하면 완성차는 인센티브를 낮게 유지하면서 원가 인플레이션을 흡수할 수 있습니다.",
      "현대차그룹의 핵심 해석은 미국 현지 생산과 HEV 라인업 확대가 관세·고정비 부담을 줄인다는 점입니다."
    ],
    questions: ["HEV 수요 확대가 BEV 재고와 완성차 인센티브에 어떤 영향을 주는가?", "현대차와 기아가 Toyota/Honda의 HEV 점유율을 가져올 수 있는 조건은 무엇인가?"],
    chart: { marketSeries: [100, 104, 108, 113, 118], demandSeries: [100, 109, 121, 136, 153] }
  } : isTelecom ? {
    oneLine: `${industry} 산업은 5G 성숙 이후 AI RAN, 6G, 주파수 경매 기대가 통신장비 투자 사이클을 다시 자극하는 상황입니다. ${scope} 범위에서는 주파수 확보 이후 기지국 투자 재개와 장비 벤더 선정이 핵심입니다.`,
    stance: "장비 중심 긍정, 통신서비스는 선별. 통신사는 배당과 요금제 변화가 주가를 지지하고, 장비사는 주파수 경매와 AI RAN 기대가 모멘텀입니다.",
    thesis: [
      "주파수 경매 이후 기지국 투자 재개 가능성이 커지면 장비 업체의 수주 기대가 먼저 반영됩니다.",
      "AI 사용량을 요금제로 묶는 시도는 통신사의 과금 단위를 데이터에서 AI 토큰/서비스로 넓히는 변화입니다.",
      "AI RAN은 6G와 별개가 아니라 5G SA 고도화와 함께 장기 설비투자 명분을 만듭니다."
    ],
    whyNow: "통신서비스의 방어적 배당 논리에 장비 투자 사이클이 결합되면 업종 내 수급이 장비와 SKT류 선도 통신사로 집중될 수 있습니다.",
    definition: "통신 산업은 무선/유선 서비스, 주파수, 기지국 장비, 네트워크 테스트, 인빌딩 솔루션, AI 네트워크 서비스를 포함합니다.",
    variables: ["주파수 경매", "5G SA/6G 투자", "AI RAN", "통신요금제", "배당/주주환원"],
    recentChange: "최근 관심은 통신사의 방어주 성격보다 AI 네트워크 투자와 주파수 경매 이후 장비 발주 가능성으로 이동하고 있습니다.",
    valueChain: [
      { stage: "Spectrum/Carrier", detail: "통신사 주파수 확보와 요금제 설계", implication: "주파수 확보는 이후 기지국 투자와 서비스 고도화의 출발점입니다." },
      { stage: "Base Station Equipment", detail: "RAN, 안테나, RF, 중계기", implication: "투자 사이클 초반에는 기지국 장비 업체가 가장 먼저 반응합니다." },
      { stage: "Testing/Optimization", detail: "망 테스트, 계측, 품질 최적화", implication: "상용망 확대 단계에서 후행 수혜가 발생합니다." },
      { stage: "AI Services", detail: "AI 요금제, 엣지 AI, AI RAN", implication: "통신사는 데이터 과금에서 AI 사용량 기반 과금으로 확장할 수 있습니다." }
    ],
    catalysts: [
      { event: "미국/국내 주파수 경매", impact: "통신사의 설비투자 기대를 높입니다.", soWhat: "초기 수혜는 통신사보다 장비 벤더와 RF/안테나 업체에 집중됩니다." },
      { event: "AI RAN/6G 논의 확대", impact: "네트워크 고도화 투자의 명분을 만듭니다.", soWhat: "장비주는 단기 실적보다 신규 투자 사이클의 옵션 가치가 먼저 반영됩니다." },
      { event: "AI 요금제 도입", impact: "통신사의 과금 단위가 데이터에서 AI 사용량으로 확장됩니다.", soWhat: "장기적으로 통신서비스의 저성장 이미지를 완화할 수 있습니다." }
    ],
    revenueCost: [
      { item: "Revenue", mechanism: "통신요금, 장비 수주, 네트워크 구축, AI 부가서비스가 매출을 만듭니다.", soWhat: "서비스는 배당 안정성, 장비는 투자 사이클 레버리지가 투자 포인트입니다." },
      { item: "Cost", mechanism: "주파수 비용, CAPEX, 장비 원가, 유지보수 비용이 부담입니다.", soWhat: "통신사는 CAPEX 부담을 요금제와 주주환원 사이에서 조절해야 합니다." },
      { item: "Margin Driver", mechanism: "DPS 성장, 장비 수주 레버리지, 벤더 선정, 해외 진출이 마진을 좌우합니다.", soWhat: "장비 업체는 특정 통신사 벤더 진입 여부가 실적의 방향을 바꿉니다." }
    ],
    keyPlayers: [
      { name: "SKT/KT/LGU+", position: "국내 통신서비스 사업자", implication: "배당과 AI 서비스 확장성이 주가 차별화 요인입니다." },
      { name: "KMW/HFR/RFHIC", position: "기지국·RF·전송 장비 업체", implication: "주파수 경매 이후 투자 사이클의 직접 수혜 후보입니다." },
      { name: "Ericsson/Samsung/Fujitsu", position: "글로벌 RAN 벤더", implication: "글로벌 통신사 벤더 선정이 국내 부품·장비 공급망으로 연결됩니다." }
    ],
    talkingPoints: [
      "통신 산업의 새로운 포인트는 배당 방어주가 아니라 AI 네트워크 투자 사이클입니다.",
      "주파수 경매는 이벤트 자체보다 이후 기지국 발주와 벤더 선정으로 연결될 때 장비주 실적이 움직입니다.",
      "AI 요금제는 통신사가 데이터 트래픽을 AI 사용량 기반 매출로 전환할 수 있다는 점에서 장기 성장 논리를 만듭니다."
    ],
    questions: ["주파수 경매 이후 어떤 장비 구간이 가장 먼저 수혜를 받는가?", "AI RAN이 통신사의 요금제와 CAPEX 논리를 어떻게 바꾸는가?"],
    chart: { marketSeries: [100, 98, 103, 116, 137], demandSeries: [100, 106, 119, 141, 168] }
  } : null;

  if (preset) {
    return {
      title: `${industry} 산업분석`,
      meta: ["기본 분석 요약", scope, period, reportType],
      analystView: {
        oneLine: preset.oneLine,
        stance: preset.stance,
        thesis: preset.thesis,
        whyNow: preset.whyNow
      },
      snapshot: {
        definition: preset.definition,
        marketVariables: preset.variables,
        recentChange: preset.recentChange
      },
      valueChain: preset.valueChain,
      catalysts: preset.catalysts,
      revenueCost: preset.revenueCost,
      keyPlayers: preset.keyPlayers,
      soWhat: {
        talkingPoints: preset.talkingPoints,
        interviewQuestions: preset.questions
      },
      chart: preset.chart,
      sourceNote: "기본 분석 요약입니다. 증권사 리포트, 공시, 통계, 뉴스 원문으로 출처를 함께 검증합니다."
    };
  }

  return {
    title: `${industry} 산업분석`,
    meta: ["기본 분석 요약", scope, period, reportType],
    analystView: {
      oneLine: `${industry} 산업은 최근 수요 성장과 비용 부담이 동시에 커지며 산업 평균보다 밸류체인별 이익 격차가 확대되는 상황입니다. ${scope} 범위에서는 최종 수요보다 가격 전가력과 공급 병목이 수익성을 가릅니다.`,
      stance: "중립적 선별. 성장성은 유효하지만 모든 플레이어가 같은 폭으로 이익을 가져가지는 못합니다.",
      thesis: [
        "최종 시장이 성장해도 공급 병목을 가진 기업은 가격과 마진을 지키고, 단순 생산자는 원가 상승을 더 크게 맞습니다.",
        "고객 락인과 장기계약이 있는 구간은 업황 둔화에도 실적 가시성이 높습니다.",
        "정책·원가·고객 투자 사이클이 같은 방향으로 움직일 때 업종 내 주가 차별화가 커집니다."
      ],
      whyNow: `${period} 흐름에서 투자자와 취업 준비생이 얻어야 할 결론은 하나입니다. 이 산업은 '성장한다'보다 '성장의 돈이 어느 구간에 남는가'가 핵심입니다.`
    },
    snapshot: {
      definition: `${industry} 산업은 최종 수요, 공급망, 가격 결정 구조, 규제/정책 변화가 연결되어 움직이는 시장입니다.`,
      marketVariables: ["최종 수요 성장률", "공급 병목과 증설 속도", "원재료/인건비/물류비", "규제와 정책 지원", "주요 고객사의 투자 사이클"],
      recentChange: "최근에는 성장 기대만으로 밸류에이션이 올라가기보다, 실제 수주와 마진 개선으로 연결되는 구간에 시장 관심이 집중되고 있습니다."
    },
    valueChain: [
      { stage: "Upstream", detail: "원재료, 핵심 부품, 기술/IP, 설비", implication: "공급 병목이 있으면 가격 결정력이 생깁니다." },
      { stage: "Production", detail: "제조, 운영, 품질, 수율, CAPEX", implication: "고정비 구조라 가동률 변화가 마진에 크게 반영됩니다." },
      { stage: "Channel", detail: "B2B 고객, 유통, 플랫폼, 장기 계약", implication: "고객 락인과 계약 구조가 실적 가시성을 만듭니다." },
      { stage: "Demand", detail: "최종 소비자/기업 수요, 교체 사이클, 투자 집행", implication: "수요가 좋아도 가격 전가가 약하면 매출만 늘고 이익은 남지 않습니다." }
    ],
    catalysts: [
      { event: "정책/규제 변화", impact: "투자 속도와 시장 진입 조건을 바꿉니다.", soWhat: "정책 수혜는 발표 직후보다 수주와 설비투자로 연결되는 기업에서 실적으로 나타납니다." },
      { event: "원가/환율/금리 변동", impact: "마진과 밸류에이션 할인율에 동시에 영향을 줍니다.", soWhat: "비용 전가력이 약한 기업은 업황 회복기에도 이익이 늦게 따라옵니다." },
      { event: "고객사 투자 사이클", impact: "B2B 산업에서는 고객의 예산 집행이 매출 성장률을 결정합니다.", soWhat: "고객의 CAPEX가 늘어나는 구간에서는 장비·부품·서비스 기업의 실적 레버리지가 먼저 나타납니다." }
    ],
    revenueCost: [
      { item: "Revenue", mechanism: "가격 x 물량 x 제품/고객 믹스로 결정됩니다.", soWhat: "성장률보다 반복 매출과 고부가 믹스 확대 여부가 중요합니다." },
      { item: "Cost", mechanism: "원재료, 인건비, 물류비, 감가상각, R&D가 핵심 비용입니다.", soWhat: "CAPEX가 큰 산업은 호황기에는 레버리지, 불황기에는 손익 부담으로 작동합니다." },
      { item: "Margin Driver", mechanism: "가격 전가력, 가동률, 수율, 규모의 경제가 마진을 좌우합니다.", soWhat: "이익률 개선이 일회성인지 구조적인지 구분해야 합니다." }
    ],
    keyPlayers: [
      { name: "국내 선도 기업군", position: "내수/수출 시장에서 규모와 고객 기반을 보유", implication: "산업 성장의 1차 수혜 후보이며, 가격 전가력이 있으면 마진 개선이 동반됩니다." },
      { name: "글로벌 리더", position: "기술, 브랜드, 원가, 고객 락인 중 하나 이상에서 우위", implication: "글로벌 리더의 투자 방향이 국내 공급망의 수주와 밸류에이션 기준점이 됩니다." },
      { name: "소부장/인프라 기업", position: "병목 구간에 위치한 공급자", implication: "최종 제품보다 높은 협상력을 가질 수 있습니다." }
    ],
    soWhat: {
      talkingPoints: [
        `${industry} 산업의 핵심은 시장 성장률이 아니라 성장의 이익이 병목 공급자와 고객 락인 기업에 집중된다는 점입니다.`,
        "원가 상승기에는 규모보다 가격 전가력, 수요 회복기에는 단순 매출보다 마진 레버리지가 기업 간 격차를 만듭니다.",
        "면접 답변에서는 '성장 산업'이라는 말보다 수혜 기업과 피해 기업이 갈리는 이유를 말하는 편이 설득력이 큽니다."
      ],
      interviewQuestions: [
        `${industry} 산업에서 가장 강한 가격 결정력을 가진 밸류체인 구간은 어디인가?`,
        "최근 뉴스 하나를 골라 매출, 비용, 마진 중 어떤 변수에 영향을 주는지 설명할 수 있는가?"
      ]
    },
    chart: {
      marketSeries: [100, 112, 126, 143, 161],
      demandSeries: [100, 116, 134, 157, 184]
    },
    sourceNote: "기본 분석 요약입니다. 증권사 리포트, 공시, 통계, 뉴스 원문으로 출처를 함께 검증합니다."
  };
}

function initDb() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS corp_master (
      corp_code TEXT PRIMARY KEY,
      corp_name TEXT NOT NULL,
      corp_eng_name TEXT,
      stock_code TEXT,
      market TEXT,
      aliases TEXT,
      normalized_name TEXT NOT NULL,
      modify_date TEXT,
      source_updated_at TEXT NOT NULL
    );

    CREATE INDEX IF NOT EXISTS idx_corp_master_stock_code ON corp_master(stock_code);
    CREATE INDEX IF NOT EXISTS idx_corp_master_normalized_name ON corp_master(normalized_name);

    CREATE TABLE IF NOT EXISTS dart_cache (
      cache_key TEXT PRIMARY KEY,
      payload TEXT NOT NULL,
      status TEXT,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS api_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      endpoint TEXT NOT NULL,
      cache_key TEXT NOT NULL,
      status TEXT,
      message TEXT,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS sync_state (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );
  `);
}

function seedCorpMaster() {
  const count = db.prepare("SELECT COUNT(*) AS count FROM corp_master").get().count;
  if (count > 0) return;
  const now = new Date().toISOString();
  const seed = [
    ["00126380", "삼성전자", "Samsung Electronics", "005930", "KOSPI", ["삼전", "Samsung Electronics"]],
    ["00164779", "SK하이닉스", "SK Hynix", "000660", "KOSPI", ["하이닉스", "SK Hynix"]],
    ["00164742", "현대자동차", "Hyundai Motor", "005380", "KOSPI", ["현대차", "Hyundai Motor"]],
    ["00106641", "기아", "Kia", "000270", "KOSPI", ["Kia", "기아자동차"]],
    ["00401731", "LG전자", "LG Electronics", "066570", "KOSPI", ["엘지전자", "LG Electronics"]],
    ["00266961", "NAVER", "NAVER", "035420", "KOSPI", ["네이버", "Naver"]],
    ["00258801", "카카오", "Kakao", "035720", "KOSPI", ["Kakao"]],
    ["00155319", "POSCO홀딩스", "POSCO Holdings", "005490", "KOSPI", ["포스코홀딩스", "POSCO", "포스코"]]
  ];
  const insert = db.prepare(`
    INSERT INTO corp_master
    (corp_code, corp_name, corp_eng_name, stock_code, market, aliases, normalized_name, modify_date, source_updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  db.exec("BEGIN");
  for (const row of seed) {
    insert.run(row[0], row[1], row[2], row[3], row[4], JSON.stringify(row[5]), normalizeQuery(row[1]), "", now);
  }
  db.exec("COMMIT");
}

async function refreshCorpMaster({ force }) {
  if (!dartKey) throw new Error("OPENDART_API_KEY is not configured.");
  const state = db.prepare("SELECT value FROM sync_state WHERE key = ?").get("corp_master_last_sync");
  const lastSync = state ? Number(state.value) : 0;
  const sevenDays = 7 * 24 * 60 * 60 * 1000;
  if (!force && Date.now() - lastSync < sevenDays) {
    return { ok: true, refreshed: false, reason: "fresh-cache" };
  }

  const url = `${DART_BASE}/corpCode.xml?crtfc_key=${encodeURIComponent(dartKey)}`;
  const response = await fetchWithTimeout(url, {}, 10000);
  if (!response.ok) throw new Error(`Failed to download corpCode.xml: HTTP ${response.status}`);
  const buffer = Buffer.from(await response.arrayBuffer());
  const zip = new AdmZip(buffer);
  const entry = zip.getEntries().find(item => item.entryName.toLowerCase().endsWith(".xml"));
  if (!entry) throw new Error("corpCode zip did not contain XML.");
  const xml = entry.getData().toString("utf8");
  const companies = parseCorpCodeXml(xml);
  upsertCompanies(companies);
  const now = new Date().toISOString();
  db.prepare(`
    INSERT INTO sync_state (key, value, updated_at)
    VALUES (?, ?, ?)
    ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
  `).run("corp_master_last_sync", String(Date.now()), now);
  return { ok: true, refreshed: true, count: companies.length, updated_at: now };
}

function parseCorpCodeXml(xml) {
  const rows = [];
  const blockRegex = /<list>([\s\S]*?)<\/list>/g;
  let match;
  while ((match = blockRegex.exec(xml))) {
    const block = match[1];
    rows.push({
      corp_code: tag(block, "corp_code"),
      corp_name: tag(block, "corp_name"),
      corp_eng_name: tag(block, "corp_eng_name"),
      stock_code: tag(block, "stock_code"),
      modify_date: tag(block, "modify_date")
    });
  }
  return rows.filter(row => row.corp_code && row.corp_name);
}

function tag(block, name) {
  const match = block.match(new RegExp(`<${name}>([\\s\\S]*?)<\\/${name}>`));
  return match ? decodeXml(match[1].trim()) : "";
}

function decodeXml(value) {
  return value
    .replaceAll("&amp;", "&")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'");
}

function upsertCompanies(companies) {
  const now = new Date().toISOString();
  const upsert = db.prepare(`
    INSERT INTO corp_master
    (corp_code, corp_name, corp_eng_name, stock_code, market, aliases, normalized_name, modify_date, source_updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(corp_code) DO UPDATE SET
      corp_name = excluded.corp_name,
      corp_eng_name = excluded.corp_eng_name,
      stock_code = excluded.stock_code,
      normalized_name = excluded.normalized_name,
      modify_date = excluded.modify_date,
      source_updated_at = excluded.source_updated_at
  `);
  db.exec("BEGIN");
  for (const company of companies) {
    upsert.run(
      company.corp_code,
      company.corp_name,
      company.corp_eng_name || "",
      company.stock_code || "",
      company.stock_code ? "LISTED" : "UNLISTED",
      "[]",
      normalizeQuery(company.corp_name),
      company.modify_date || "",
      now
    );
  }
  db.exec("COMMIT");
}

function resolveCompanies(query) {
  const raw = String(query || "").trim();
  const normalized = normalizeQuery(raw);
  if (!normalized) {
    return db.prepare("SELECT * FROM corp_master WHERE stock_code != '' ORDER BY corp_name LIMIT 8").all().map(publicCorp);
  }

  const exact = db.prepare(`
    SELECT *, 100 AS score FROM corp_master
    WHERE corp_code = ? OR stock_code = ? OR normalized_name = ?
    LIMIT 8
  `).all(raw, raw, normalized);
  const seen = new Set(exact.map(row => row.corp_code));

  const fuzzy = db.prepare(`
    SELECT *, 60 AS score FROM corp_master
    WHERE normalized_name LIKE ? OR corp_name LIKE ? OR aliases LIKE ?
    ORDER BY CASE WHEN stock_code != '' THEN 0 ELSE 1 END, corp_name
    LIMIT 12
  `).all(`%${normalized}%`, `%${raw}%`, `%${raw}%`).filter(row => !seen.has(row.corp_code));

  return [...exact, ...fuzzy].slice(0, 8).map(publicCorp);
}

function publicCorp(row) {
  return {
    corp_code: row.corp_code,
    name: row.corp_name,
    corp_name: row.corp_name,
    corp_eng_name: row.corp_eng_name || "",
    stock_code: row.stock_code || "",
    market: row.market || (row.stock_code ? "LISTED" : "UNLISTED"),
    aliases: safeJson(row.aliases, []),
    modify_date: row.modify_date || ""
  };
}

async function buildFinancialsResponse({ corpCode, baseYear, years, reportCode, fsDiv }) {
  const company = db.prepare("SELECT * FROM corp_master WHERE corp_code = ?").get(corpCode) || null;
  const periodYears = Array.from({ length: years }, (_, index) => baseYear - years + 1 + index);
  const periods = await Promise.all(periodYears.map(year =>
    getFinancialPeriod({ corpCode, year, reportCode, fsDiv }).catch(error => ({
      year,
      report_code: reportCode,
      fs_div: fsDiv,
      status: "timeout",
      message: error.message,
      raw_count: 0,
      raw_items: []
    }))
  ));

  const normalized = periods.map(period => normalizeFinancialPeriod(period));
  addDerivedMetrics(normalized);
  const quality = buildQuality(normalized, fsDiv);
  const latest = [...normalized].reverse().find(period => period.metrics && Object.keys(period.metrics).length) || null;
  const cashFlowPattern = latest ? classifyCashFlowPattern(latest.metrics) : null;
  const publicCompany = company ? publicCorp(company) : { corp_code: corpCode, name: corpCode };
  const marketData = await buildMarketData({ company: publicCompany, corpCode, baseYear, reportCode });
  const valuation = buildValuation(latest ? latest.metrics : {}, marketData);

  return {
    ok: true,
    company: publicCompany,
    request: { corp_code: corpCode, base_year: baseYear, years, report_code: reportCode, fs_div: fsDiv },
    periods: normalized,
    quality,
    market_data: marketData,
    valuation,
    cash_flow_pattern: cashFlowPattern,
    insights: latest ? buildCompanyInsights(publicCompany, latest, { quality, cash_flow_pattern: cashFlowPattern, market_data: marketData, valuation }) : [],
    lineage: {
      source: "OpenDART",
      endpoint: "fnlttSinglAcntAll",
      cached: true,
      generated_at: new Date().toISOString()
    }
  };
}

async function buildMarketData({ company, corpCode, baseYear, reportCode }) {
  const marketData = {
    no_of_shares: null,
    stock_price: null,
    market_cap: null,
    currency: "KRW",
    price_date: null,
    sources: [],
    warnings: []
  };

  if (!company || !company.stock_code) {
    marketData.warnings.push("비상장 또는 종목코드 미확보로 주가 리서치를 건너뜁니다.");
    return marketData;
  }

  const [sharesResult, priceResult] = await Promise.allSettled([
    fetchDartNoOfShares({ corpCode, baseYear, reportCode }),
    fetchListedStockPrice(company)
  ]);

  if (sharesResult.status === "fulfilled" && sharesResult.value) {
    marketData.no_of_shares = sharesResult.value.no_of_shares;
    marketData.sources.push(sharesResult.value.source);
  } else {
    marketData.warnings.push("OpenDART 발행주식수 자동 조회 실패. Summary의 No. of Shares 셀은 AI/수동 리서치 보완 대상입니다.");
  }

  if (priceResult.status === "fulfilled" && priceResult.value) {
    marketData.stock_price = priceResult.value.stock_price;
    marketData.price_date = priceResult.value.price_date;
    marketData.sources.push(priceResult.value.source);
  } else {
    marketData.warnings.push("상장 시세 자동 조회 실패. Summary의 Stock Price 셀은 AI/수동 리서치 보완 대상입니다.");
  }

  if (Number.isFinite(marketData.no_of_shares) && Number.isFinite(marketData.stock_price)) {
    marketData.market_cap = marketData.no_of_shares * marketData.stock_price;
  }

  return marketData;
}

async function fetchDartNoOfShares({ corpCode, baseYear, reportCode }) {
  if (!dartKey || !/^\d{8}$/.test(String(corpCode || ""))) return null;
  const params = {
    corp_code: corpCode,
    bsns_year: String(baseYear),
    reprt_code: reportCode
  };
  const cacheKey = `stockTotqySttus:${new URLSearchParams(params).toString()}`;
  const payload = await dartGet("/stockTotqySttus.json", params, cacheKey);
  const rows = Array.isArray(payload.list) ? payload.list : [];
  if (!rows.length) return null;
  const clean = value => String(value || "").replace(/\s+/g, "");
  const common =
    rows.find(row => clean(row.se).includes("보통")) ||
    rows.find(row => clean(row.se).includes("합계")) ||
    rows[0];
  const shares = parseAmount(common.istc_totqy || common.distb_stock_co || common.stock_totqy);
  if (!Number.isFinite(shares)) return null;
  return {
    no_of_shares: shares,
    source: {
      type: "shares",
      name: "OpenDART 주식의 총수 현황",
      endpoint: "stockTotqySttus",
      basis: `${baseYear} ${reportCode}`,
      label: common.se || "발행주식총수"
    }
  };
}

async function fetchListedStockPrice(company) {
  const stockCode = String(company.stock_code || "").trim();
  if (!/^\d{6}$/.test(stockCode)) return null;
  const suffixes = String(company.market || "").toUpperCase().includes("KOSDAQ") ? ["KQ", "KS"] : ["KS", "KQ"];
  for (const suffix of suffixes) {
    const symbol = `${stockCode}.${suffix}`;
    const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?range=5d&interval=1d`;
    try {
      const response = await fetchWithTimeout(url, {
        headers: {
          "User-Agent": "Mozilla/5.0 BIT-Analysis/1.0"
        }
      }, 5000);
      if (!response.ok) continue;
      const payload = await response.json();
      const result = payload && payload.chart && payload.chart.result && payload.chart.result[0];
      if (!result) continue;
      const quote = result.indicators && result.indicators.quote && result.indicators.quote[0];
      const closes = quote && Array.isArray(quote.close) ? quote.close.filter(Number.isFinite) : [];
      const price = Number(result.meta && result.meta.regularMarketPrice) || closes.at(-1);
      if (!Number.isFinite(price)) continue;
      const timestamps = Array.isArray(result.timestamp) ? result.timestamp : [];
      const latestTimestamp = timestamps.at(-1);
      const priceDate = latestTimestamp ? new Date(latestTimestamp * 1000).toISOString().slice(0, 10) : new Date().toISOString().slice(0, 10);
      return {
        stock_price: Math.round(price),
        price_date: priceDate,
        source: {
          type: "stock_price",
          name: "Yahoo Finance chart",
          symbol,
          url,
          basis: priceDate
        }
      };
    } catch (error) {
      // Try the next market suffix before surfacing a warning.
    }
  }
  return null;
}

function buildValuation(metrics, marketData) {
  const marketCap = Number.isFinite(marketData.market_cap) ? marketData.market_cap : null;
  const netDebt = Number.isFinite(metrics.net_debt) ? metrics.net_debt : null;
  const netIncome = Number.isFinite(metrics.net_income) ? metrics.net_income : null;
  const ebitda = Number.isFinite(metrics.ebitda) ? metrics.ebitda : null;
  const ev = marketCap !== null && netDebt !== null ? marketCap + netDebt : null;
  return {
    market_cap: marketCap,
    per: marketCap !== null && netIncome ? marketCap / netIncome : null,
    enterprise_value: ev,
    ev_ebitda: ev !== null && ebitda ? ev / ebitda : null
  };
}

async function getFinancialPeriod({ corpCode, year, reportCode, fsDiv }) {
  let result = await fetchFinancial({ corpCode, year, reportCode, fsDiv });
  if ((!result.list || !result.list.length) && fsDiv === "CFS") {
    result = await fetchFinancial({ corpCode, year, reportCode, fsDiv: "OFS" });
  }
  return {
    year,
    report_code: reportCode,
    fs_div: result.request_fs_div,
    status: result.status,
    message: result.message,
    raw_count: Array.isArray(result.list) ? result.list.length : 0,
    raw_items: Array.isArray(result.list) ? result.list : []
  };
}

async function fetchFinancial({ corpCode, year, reportCode, fsDiv }) {
  const params = {
    corp_code: corpCode,
    bsns_year: String(year),
    reprt_code: reportCode,
    fs_div: fsDiv
  };
  const cacheKey = `fnlttSinglAcntAll:${new URLSearchParams(params).toString()}`;
  const cached = db.prepare("SELECT payload FROM dart_cache WHERE cache_key = ?").get(cacheKey);
  if (cached) return { ...JSON.parse(cached.payload), request_fs_div: fsDiv };

  const payload = await dartGet("/fnlttSinglAcntAll.json", params, cacheKey);
  db.prepare(`
    INSERT INTO dart_cache (cache_key, payload, status, created_at)
    VALUES (?, ?, ?, ?)
    ON CONFLICT(cache_key) DO UPDATE SET payload = excluded.payload, status = excluded.status, created_at = excluded.created_at
  `).run(cacheKey, JSON.stringify(payload), payload.status || "", new Date().toISOString());
  return { ...payload, request_fs_div: fsDiv };
}

async function dartGet(endpoint, params, cacheKey) {
  const url = new URL(`${DART_BASE}${endpoint}`);
  url.searchParams.set("crtfc_key", dartKey);
  for (const [key, value] of Object.entries(params)) url.searchParams.set(key, value);
  const response = await fetchWithTimeout(url, {}, 12000);
  if (!response.ok) throw new Error(`OpenDART HTTP ${response.status}`);
  const payload = await response.json();
  db.prepare("INSERT INTO api_logs (endpoint, cache_key, status, message, created_at) VALUES (?, ?, ?, ?, ?)")
    .run(endpoint, cacheKey, payload.status || "", payload.message || "", new Date().toISOString());
  if (payload.status && !["000", "013"].includes(payload.status)) {
    throw new Error(`OpenDART ${payload.status}: ${payload.message || "request failed"}`);
  }
  return payload;
}

function normalizeFinancialPeriod(period) {
  const metrics = {};
  const evidence = {};
  const rules = [
    ["revenue", ["매출액", "수익(매출액)", "영업수익", "매출"], ["매출원가", "매출채권", "매출총이익"]],
    ["operating_income", ["영업이익"], []],
    ["depreciation", ["감가상각비"], []],
    ["amortization", ["무형자산상각비"], []],
    ["net_income", ["당기순이익", "분기순이익", "반기순이익"], ["총포괄"]],
    ["total_assets", ["자산총계"], []],
    ["total_liabilities", ["부채총계"], []],
    ["total_equity", ["자본총계"], []],
    ["short_term_debt", ["단기차입금", "단기금융부채"], []],
    ["long_term_debt", ["장기차입금", "장기금융부채"], []],
    ["cash_and_equivalents", ["현금및현금성자산", "현금 및 현금성자산"], []],
    ["operating_cash_flow", ["영업활동현금흐름", "영업활동으로 인한 현금흐름"], []],
    ["investing_cash_flow", ["투자활동현금흐름", "투자활동으로 인한 현금흐름"], []],
    ["financing_cash_flow", ["재무활동현금흐름", "재무활동으로 인한 현금흐름"], []]
  ];

  for (const [metric, includes, excludes] of rules) {
    const item = findFinancialItem(period.raw_items, includes, excludes);
    if (!item) continue;
    const amount = parseAmount(item.thstrm_amount || item.amount);
    if (amount === null) continue;
    metrics[metric] = amount;
    evidence[metric] = {
      account_name: item.account_nm,
      statement: item.sj_nm || item.sj_div || "",
      currency: item.currency || "KRW"
    };
  }

  if (metrics.revenue && metrics.operating_income !== undefined) {
    metrics.operating_margin = metrics.operating_income / metrics.revenue;
  }
  if (metrics.total_equity && metrics.total_liabilities !== undefined) {
    metrics.debt_ratio = metrics.total_liabilities / metrics.total_equity;
  }
  if (metrics.operating_income !== undefined && (metrics.depreciation !== undefined || metrics.amortization !== undefined)) {
    metrics.ebitda = metrics.operating_income + (metrics.depreciation || 0) + (metrics.amortization || 0);
  }
  if (metrics.short_term_debt !== undefined || metrics.long_term_debt !== undefined) {
    metrics.total_debt = (metrics.short_term_debt || 0) + (metrics.long_term_debt || 0);
  }
  if (metrics.total_debt !== undefined && metrics.cash_and_equivalents !== undefined) {
    metrics.net_debt = metrics.total_debt - metrics.cash_and_equivalents;
  }
  metrics.cash_flow_pattern = classifyCashFlowPattern(metrics);

  return {
    year: period.year,
    report_code: period.report_code,
    fs_div: period.fs_div,
    status: period.status,
    message: period.message,
    raw_count: period.raw_count,
    metrics,
    evidence
  };
}

function findFinancialItem(items, includes, excludes) {
  const clean = value => String(value || "").replace(/\s+/g, "");
  const candidates = items.filter(item => {
    const name = clean(item.account_nm);
    return includes.some(word => name.includes(clean(word))) && !excludes.some(word => name.includes(clean(word)));
  });
  return candidates.sort((a, b) => {
    const statementScore = item => clean(item.sj_nm).includes("손익계산서") || clean(item.sj_nm).includes("포괄손익계산서") ? 0 : 1;
    const exactScore = item => includes.some(word => clean(item.account_nm) === clean(word)) ? 0 : 1;
    return statementScore(a) - statementScore(b) || exactScore(a) - exactScore(b) || clean(a.account_nm).length - clean(b.account_nm).length;
  })[0];
}

function parseAmount(value) {
  if (value === undefined || value === null) return null;
  const text = String(value).trim();
  if (!text || text === "-") return null;
  const negative = /^\(.*\)$/.test(text);
  const numeric = Number(text.replace(/[(),]/g, ""));
  if (!Number.isFinite(numeric)) return null;
  return negative ? -numeric : numeric;
}

function addDerivedMetrics(periods) {
  for (let index = 1; index < periods.length; index += 1) {
    const current = periods[index].metrics;
    const previous = periods[index - 1].metrics;
    if (current.revenue !== undefined && previous.revenue) {
      current.revenue_growth = current.revenue / previous.revenue - 1;
    }
  }
}

function buildQuality(periods, requestedFsDiv) {
  const latest = [...periods].reverse().find(period => period.metrics && Object.keys(period.metrics).length) || null;
  const latestMetricCount = latest ? Object.keys(latest.metrics).filter(key => !key.endsWith("_growth") && !key.endsWith("_margin") && key !== "debt_ratio" && key !== "cash_flow_pattern").length : 0;
  const availableYears = periods.filter(period => Object.keys(period.metrics).length).length;
  const hasCfs = periods.some(period => period.fs_div === "CFS" && Object.keys(period.metrics).length);
  let score = "low";
  if (latestMetricCount >= 6 && availableYears >= 3 && (requestedFsDiv !== "CFS" || hasCfs)) score = "high";
  else if (latestMetricCount >= 4 && availableYears >= 2) score = "medium";
  return {
    score,
    latest_year: latest ? latest.year : null,
    available_years: availableYears,
    latest_metric_count: latestMetricCount,
    checks: [
      hasCfs ? "연결 기준 데이터 확인" : "연결 기준 미확인 또는 별도 기준 fallback",
      `${availableYears}개 연도 데이터 확보`,
      `${latestMetricCount}개 핵심 계정 매핑`
    ]
  };
}

function classifyCashFlowPattern(metrics) {
  const op = sign(metrics.operating_cash_flow);
  const inv = sign(metrics.investing_cash_flow);
  const fin = sign(metrics.financing_cash_flow);
  if (!op || !inv || !fin) {
    return {
      type: "판단 보류",
      signs: { operating: op, investing: inv, financing: fin },
      description: "영업/투자/재무 현금흐름 중 일부가 없어 패턴 분류를 보류합니다."
    };
  }

  const key = `${op}${inv}${fin}`;
  const patterns = {
    "-+-": ["스타트업", "사업은 아직 돈을 벌지 못하고 외부 자금으로 버티며 투자하는 구간입니다."],
    "+-+": ["성장기업", "영업에서 돈을 벌기 시작했지만 CAPEX와 확장 투자 때문에 추가 자금 조달이 필요한 구간입니다."],
    "+--": ["우량기업", "영업현금흐름으로 투자와 차입 상환 또는 배당을 감당하는 안정적 패턴입니다."],
    "++-": ["전환기업", "현금 유입은 있으나 투자 회수나 자산 매각 성격이 섞인 패턴입니다."],
    "++-alt": ["성숙기업", "신규 투자보다 회수와 주주환원 또는 부채 상환이 중심인 구간입니다."],
    "---": ["쇠퇴기업", "영업, 투자, 재무 현금흐름이 모두 유출로 나타나면 현금 소진과 지속 가능성 점검이 필요합니다."],
    "-++": ["부실기업", "영업은 부진하고 투자 회수와 외부 조달로 버티는 패턴입니다."],
    "-+-alt": ["정리기업", "사업 정리와 재무구조 재편 과정일 수 있어 일회성 현금흐름을 분리해야 합니다."]
  };
  const selected = patterns[key] || ["혼합형", "사업 이벤트, 일회성 투자, 차입 상환이 함께 섞인 패턴입니다."];
  return {
    type: selected[0],
    signs: { operating: op, investing: inv, financing: fin },
    description: selected[1]
  };
}

function sign(value) {
  if (!Number.isFinite(value) || value === 0) return "0";
  return value > 0 ? "+" : "-";
}

function buildCompanyInsights(company, latest, financials) {
  const metrics = latest.metrics;
  const pattern = financials.cash_flow_pattern || metrics.cash_flow_pattern || classifyCashFlowPattern(metrics);
  const valuation = financials.valuation || buildValuation(metrics, financials.market_data || {});
  const marketLine = Number.isFinite(valuation.per) && Number.isFinite(valuation.ev_ebitda)
    ? `발행주식수와 주가 리서치 기준 PER은 ${valuation.per.toFixed(1)}배, EV/EBITDA는 ${valuation.ev_ebitda.toFixed(1)}배입니다.`
    : "발행주식수와 주가는 별도 시장 데이터 리서치 후 PER, EV/EBITDA를 자동 계산합니다.";
  return [
    `${latest.year}년 DART 기준 매출은 ${formatWon(metrics.revenue)}, EBIT는 ${formatWon(metrics.operating_income)}, EBITDA는 ${formatWon(metrics.ebitda)}입니다.`,
    `순차입금은 ${formatWon(metrics.net_debt)}, 부채비율은 ${formatPercent(metrics.debt_ratio)}로 재무 안정성 판단의 1차 기준이 됩니다.`,
    `현금흐름 패턴은 ${pattern.type}입니다. ${pattern.description}`,
    Number.isFinite(metrics.revenue_growth)
      ? `최근 성장성은 매출 증가율 ${formatPercent(metrics.revenue_growth)}를 기준으로 산업 사이클과 회사 고유 경쟁력을 분리해 읽는 구조입니다.`
      : "전년 대비 성장률은 비교 연도 데이터가 확보되면 자동 계산됩니다.",
    marketLine,
    `${company.name || "해당 기업"} 분석 산출물에는 숫자 자체보다 고객(C), 상품(P), 채널(C), 재무 패턴을 연결한 해석이 포함되어야 합니다.`
  ];
}

async function buildCompanyWorkbook(financials) {
  if (fs.existsSync(companyTemplatePath)) {
    return buildCompanyWorkbookFromProvidedTemplate(financials);
  }
  return buildGeneratedCompanyWorkbook(financials);
}

async function buildGeneratedCompanyWorkbook(financials) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "BIT Analysis";
  workbook.created = new Date();
  workbook.modified = new Date();

  const summary = workbook.addWorksheet("Summary", { views: [{ state: "frozen", ySplit: 4 }] });
  const templateSheet = workbook.addWorksheet("1", { views: [{ state: "frozen", ySplit: 5 }] });
  const financialSheet = workbook.addWorksheet("DART Financials", { views: [{ state: "frozen", ySplit: 1 }] });
  const marketSheet = workbook.addWorksheet("Market Data");
  const insights = workbook.addWorksheet("Insights");
  const lineage = workbook.addWorksheet("Lineage");

  buildTemplateInputSheet(templateSheet, financials);
  buildTemplateSummarySheet(summary, financials);
  buildFinancialSheet(financialSheet, financials);
  buildMarketDataSheet(marketSheet, financials);
  buildInsightsSheet(insights, financials);
  buildLineageSheet(lineage, financials);

  for (const sheet of workbook.worksheets) {
    sheet.properties.defaultRowHeight = 22;
    sheet.eachRow(row => {
      row.alignment = { vertical: "top", wrapText: true };
    });
  }

  return workbook.xlsx.writeBuffer();
}

async function buildCompanyWorkbookFromProvidedTemplate(financials) {
  const zip = new AdmZip(companyTemplatePath);
  const summaryPath = worksheetPath(zip, "Summary");
  const inputPath = worksheetPath(zip, "1");
  if (!summaryPath || !inputPath) {
    return buildGeneratedCompanyWorkbook(financials);
  }

  zip.updateFile(inputPath, Buffer.from(fillProvidedTemplateInputXml(zip.readAsText(inputPath), financials)));
  zip.updateFile(summaryPath, Buffer.from(fillProvidedTemplateSummaryXml(zip.readAsText(summaryPath), financials)));

  const workbookXml = zip.readAsText("xl/workbook.xml");
  const nextWorkbookXml = setWorkbookFullRecalc(hideWorkbookSheets(workbookXml, ["2", "3"]));
  zip.updateFile("xl/workbook.xml", Buffer.from(nextWorkbookXml));
  return zip.toBuffer();
}

function worksheetPath(zip, sheetName) {
  const workbookXml = zip.readAsText("xl/workbook.xml");
  const relsXml = zip.readAsText("xl/_rels/workbook.xml.rels");
  const sheetTag = workbookXml.match(new RegExp(`<sheet\\b(?=[^>]*\\bname="${escapeRegExp(sheetName)}")[^>]*>`, "u"))?.[0];
  const relId = sheetTag?.match(/\br:id="([^"]+)"/u)?.[1];
  if (!relId) return "";
  const relTag = relsXml.match(new RegExp(`<Relationship\\b(?=[^>]*\\bId="${escapeRegExp(relId)}")[^>]*>`, "u"))?.[0];
  const target = relTag?.match(/\bTarget="([^"]+)"/u)?.[1];
  if (!target) return "";
  return target.startsWith("/") ? target.replace(/^\//u, "") : `xl/${target}`.replace(/\/[^/]+\/\.\.\//gu, "/");
}

function fillProvidedTemplateInputXml(xml, financials) {
  const periods = templatePeriods(financials);
  const cells = {
    A3: textCell("회사명"),
    B3: textCell(financials.company.name || financials.company.corp_code),
    C3: textCell("상장여부"),
    D3: textCell(financials.company.stock_code ? "상장" : "비상장"),
    G3: textCell("(단위: 백만, DART 원천값/1,000,000)")
  };

  for (let index = 0; index < 5; index += 1) {
    const col = String.fromCharCode("B".charCodeAt(0) + index);
    const period = periods[index];
    cells[`${col}5`] = period ? numberCell(Number(period.year)) : blankCell();
  }

  const rows = [
    [6, "revenue"],
    [7, "operating_income"],
    [8, "depreciation"],
    [9, "amortization"],
    [11, "net_income"],
    [12, "total_assets"],
    [13, "total_liabilities"],
    [14, "total_equity"],
    [15, "short_term_debt"],
    [16, "long_term_debt"],
    [18, "cash_and_equivalents"],
    [20, "operating_cash_flow"],
    [21, "investing_cash_flow"],
    [22, "financing_cash_flow"]
  ];

  for (let index = 0; index < 5; index += 1) {
    const col = String.fromCharCode("B".charCodeAt(0) + index);
    const period = periods[index];
    for (const [row, key] of rows) {
      cells[`${col}${row}`] = period ? numberOrBlankCell(toMillion(period.metrics[key])) : blankCell();
    }
    cells[`${col}10`] = period && Number.isFinite(period.metrics.ebitda)
      ? numberCell(toMillion(period.metrics.ebitda))
      : formulaCell(`IF(OR(ISBLANK(${col}8),ISBLANK(${col}9)),"D&A 보강 필요",${col}7+${col}8+${col}9)`);
    cells[`${col}17`] = period ? formulaCell(`${col}15+${col}16`) : blankCell();
    cells[`${col}19`] = period ? formulaCell(`${col}17-${col}18`) : blankCell();
  }

  return updateSheetCells(xml, cells);
}

function fillProvidedTemplateSummaryXml(xml, financials) {
  const periods = templatePeriods(financials);
  const latest = latestPeriod(financials);
  const metrics = latest.metrics || {};
  const marketData = financials.market_data || {};
  const valuation = financials.valuation || buildValuation(metrics, marketData);
  const pattern = financials.cash_flow_pattern || metrics.cash_flow_pattern || classifyCashFlowPattern(metrics);
  const company = financials.company.name || financials.company.corp_code;
  const cells = {
    B3: textCell(new Date().toISOString().slice(0, 10)),
    C5: numberCell(1),
    L6: numberOrTextCell(marketData.no_of_shares, "리서치 필요"),
    M6: numberOrTextCell(marketData.stock_price, "리서치 필요"),
    N6: formulaCell('IFERROR(L6*M6/1000000,"입력 필요")'),
    N7: formulaCell('IFERROR(N6/I8,"-")'),
    N8: formulaCell('IFERROR(N6+I17,"-")'),
    N9: formulaCell('IFERROR(N8/I7,"-")'),
    O17: textCell(`○ 영업현금흐름: ${formatWon(metrics.operating_cash_flow)}. ${pattern.type}: ${pattern.description}`),
    O19: textCell(`○ 투자현금흐름: ${formatWon(metrics.investing_cash_flow)}. 설비투자, M&A, 금융상품 운용 여부를 분리 확인`),
    O21: textCell(`○ 재무현금흐름: ${formatWon(metrics.financing_cash_flow)}. 차입, 상환, 배당, 증자 이벤트 확인`),
    Y5: textCell(`○ ${company} 기업 기본정보와 사업 포지션`),
    Z5: textCell("○ 사업적 강점"),
    AA5: textCell("DART API 기준 자동 생성. 주가/발행주식수는 별도 시장 데이터 기준일 확인 필요."),
    Y14: textCell("○ 기업 분석 메모"),
    Y15: textCell("사업보고서의 사업부, 주요 제품, 고객, 시장 전망과 재무 변화 연결"),
    Y16: textCell("매출 성장, 마진, 현금흐름으로 확인되는 회사 고유 경쟁력을 우선 해석"),
    Z14: textCell("○ 리스크 요인"),
    Z15: textCell("매출 둔화, 마진 하락, 운전자본 부담, 차입금 증가, 투자회수 지연"),
    AA14: textCell("작성 가이드 반영"),
    AA15: textCell(`DART 최신 사업보고서 우선, 5개년 확보 시 과거 공시 조회, 음수 부호 유지, D&A 주석 확인. ${formatMarketDataNote(marketData, valuation)}`)
  };

  for (let index = 0; index < 5; index += 1) {
    const col = String.fromCharCode("E".charCodeAt(0) + index);
    const period = periods[index];
    cells[`${col}4`] = period ? numberCell(Number(period.year)) : blankCell();
  }

  return hideSheetRows(updateSheetCells(xml, cells), 22, 98);
}

function updateSheetCells(xml, cells) {
  let nextXml = xml;
  for (const [ref, cell] of Object.entries(cells)) {
    nextXml = updateSheetCell(nextXml, ref, cell);
  }
  return nextXml;
}

function updateSheetCell(xml, ref, cell) {
  const existingCell = xml.match(new RegExp(`<c\\b(?=[^>]*\\br="${escapeRegExp(ref)}")[^>]*\\/>|<c\\b(?=[^>]*\\br="${escapeRegExp(ref)}")[\\s\\S]*?<\\/c>`, "u"))?.[0];
  const style = existingCell?.match(/\bs="([^"]+)"/u)?.[1];
  const nextCell = cellXml(ref, cell, style);
  if (existingCell) return xml.replace(existingCell, nextCell);

  const rowNumber = cellRef(ref).row;
  const rowRegex = new RegExp(`(<row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>)([\\s\\S]*?)(<\\/row>)`, "u");
  if (rowRegex.test(xml)) {
    return xml.replace(rowRegex, `$1$2${nextCell}$3`);
  }
  return xml.replace("</sheetData>", `<row r="${rowNumber}">${nextCell}</row></sheetData>`);
}

function cellXml(ref, cell, style) {
  const styleAttr = style ? ` s="${escapeXml(style)}"` : "";
  if (cell.type === "formula") {
    return `<c r="${ref}"${styleAttr}><f>${escapeXml(cell.formula)}</f></c>`;
  }
  if (cell.type === "blank" || cell.value === null || cell.value === undefined || cell.value === "") {
    return `<c r="${ref}"${styleAttr}/>`;
  }
  if (cell.type === "number") {
    return `<c r="${ref}"${styleAttr}><v>${cell.value}</v></c>`;
  }
  return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t>${escapeXml(String(cell.value))}</t></is></c>`;
}

function textCell(value) {
  return { type: "text", value };
}

function numberCell(value) {
  return { type: "number", value };
}

function numberOrBlankCell(value) {
  return Number.isFinite(value) ? numberCell(value) : blankCell();
}

function numberOrTextCell(value, fallback) {
  return Number.isFinite(value) ? numberCell(value) : textCell(fallback);
}

function formulaCell(formula) {
  return { type: "formula", formula };
}

function blankCell() {
  return { type: "blank" };
}

function cellRef(ref) {
  const match = String(ref).match(/^([A-Z]+)(\d+)$/u);
  return { col: match?.[1] || "", row: Number(match?.[2] || 0) };
}

function hideSheetRows(xml, start, end) {
  let nextXml = xml;
  for (let row = start; row <= end; row += 1) {
    nextXml = nextXml.replace(new RegExp(`<row\\b(?=[^>]*\\br="${row}")(?![^>]*\\bhidden=)[^>]*>`, "u"), match => match.replace(/>$/u, ' hidden="1">'));
  }
  return nextXml;
}

function hideWorkbookSheets(workbookXml, sheetNames) {
  return sheetNames.reduce((xml, sheetName) => xml.replace(
    new RegExp(`<sheet\\b(?=[^>]*\\bname="${escapeRegExp(sheetName)}")(?![^>]*\\bstate=)[^>]*>`, "u"),
    match => match.endsWith("/>")
      ? match.replace(/\/>$/u, ' state="hidden"/>')
      : match.replace(/>$/u, ' state="hidden">')
  ), workbookXml);
}

function setWorkbookFullRecalc(workbookXml) {
  const calcPr = '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>';
  if (/<calcPr\b[^>]*\/>/u.test(workbookXml)) return workbookXml.replace(/<calcPr\b[^>]*\/>/u, calcPr);
  if (/<calcPr\b[\s\S]*?<\/calcPr>/u.test(workbookXml)) return workbookXml.replace(/<calcPr\b[\s\S]*?<\/calcPr>/u, calcPr);
  return workbookXml.replace("</workbook>", `${calcPr}</workbook>`);
}

function escapeXml(value) {
  return String(value)
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;")
    .replace(/"/gu, "&quot;")
    .replace(/'/gu, "&apos;");
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}

function latestPeriod(financials) {
  return [...financials.periods].reverse().find(period => period.metrics && Object.keys(period.metrics).length) || { metrics: {} };
}

function templatePeriods(financials) {
  const periods = financials.periods.filter(period => period.metrics && Object.keys(period.metrics).length);
  return periods.slice(-5);
}

function buildTemplateInputSheet(sheet, financials) {
  const periods = templatePeriods(financials);
  sheet.columns = [
    { width: 28 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 },
    { width: 28 }, { width: 4 }, { width: 34 }, { width: 28 }, { width: 28 }, { width: 28 }
  ];
  sheet.getCell("D2").value = "코스피, 코스닥, 코넥스, 비상장";
  sheet.getCell("A3").value = "회사명";
  sheet.getCell("B3").value = financials.company.name || financials.company.corp_code;
  sheet.getCell("C3").value = "상장여부";
  sheet.getCell("D3").value = financials.company.stock_code ? "상장" : "비상장";
  sheet.getCell("G3").value = "(단위: 백만, DART 원천값/1,000,000)";
  sheet.getRow(5).values = [null, null, ...periods.map(period => period.year), "data"];

  const rows = [
    [6, "매출액(수익)", "revenue", "손익계산서"],
    [7, "영업이익(EBIT)", "operating_income", "손익계산서"],
    [8, "감가상각비", "depreciation", "현금흐름표 or 현금흐름표(주석)"],
    [9, "무형자산상각비", "amortization", "현금흐름표 or 현금흐름표(주석)"],
    [10, "EBITDA", "ebitda", "자동계산"],
    [11, "당기순이익", "net_income", "손익계산서"],
    [12, "자산총계", "total_assets", "재무상태표"],
    [13, "부채총계", "total_liabilities", "재무상태표"],
    [14, "자본총계", "total_equity", "재무상태표"],
    [15, "단기차입금", "short_term_debt", "재무상태표"],
    [16, "장기차입금", "long_term_debt", "재무상태표"],
    [17, "총차입금", "total_debt", "자동계산"],
    [18, "현금성자산", "cash_and_equivalents", "재무상태표"],
    [19, "순차입금(net debt)", "net_debt", "자동계산"],
    [20, "영업활동현금흐름 (cash flows from operations)", "operating_cash_flow", "현금흐름표"],
    [21, "투자활동현금흐름 (cash flows from investing)", "investing_cash_flow", "현금흐름표"],
    [22, "재무활동현금흐름 (cash flows from financing)", "financing_cash_flow", "현금흐름표"]
  ];

  for (const [row, label, key, source] of rows) {
    sheet.getCell(row, 1).value = label;
    sheet.getCell(row, 7).value = source;
    periods.forEach((period, index) => {
      const col = 2 + index;
      const addr = sheet.getCell(row, col).address;
      if (row === 10) {
        const ebit = sheet.getCell(7, col).address;
        const depreciation = sheet.getCell(8, col).address;
        const amortization = sheet.getCell(9, col).address;
        sheet.getCell(addr).value = { formula: `IF(OR(ISBLANK(${depreciation}),ISBLANK(${amortization})),"리서치 필요",${ebit}+${depreciation}+${amortization})` };
      }
      else if (row === 17) sheet.getCell(addr).value = { formula: `${sheet.getCell(15, col).address}+${sheet.getCell(16, col).address}` };
      else if (row === 19) sheet.getCell(addr).value = { formula: `${sheet.getCell(17, col).address}-${sheet.getCell(18, col).address}` };
      else sheet.getCell(addr).value = toMillion(period.metrics[key]);
    });
  }

  sheet.getCell("I6").value = "손익계산서";
  sheet.getCell("J6").value = "재무상태표";
  sheet.getCell("K6").value = "현금흐름표";
  sheet.getCell("I7").value = "참조";
  sheet.getCell("K7").value = "주석 참조";
  sheet.getCell("I10").value = "EBITDA = 영업이익 + 감가상각비 + 무형자산상각비";
  sheet.getCell("I17").value = "총차입금 = 단기차입금 + 장기차입금";
  sheet.getCell("I19").value = "순차입금 = 총차입금 - 현금성자산";
  sheet.getCell("A24").value = "* Summary의 No. of Shares는 OpenDART 주식의 총수 현황, Stock Price는 상장 시세 리서치 기준으로 자동 보완합니다.";
  sheet.getCell("A25").value = "** template의 단위는 백만. DART API 원천값(원)은 입력 시 /1,000,000으로 변환합니다.";
  sheet.getCell("A26").value = "*** 감가상각비/무형자산상각비는 현금흐름표 또는 주석에서 확인합니다.";
  sheet.getCell("A27").value = "**** DART 최신 사업보고서 숫자를 우선하며, 5개년 확보를 위해 과거 사업보고서를 함께 조회합니다.";

  sheet.getRange?.("A1:L27");
  styleTemplateSheet(sheet);
}

function buildTemplateSummarySheet(sheet, financials) {
  const periods = templatePeriods(financials);
  const latest = latestPeriod(financials);
  const metrics = latest.metrics || {};
  const marketData = financials.market_data || {};
  const valuation = financials.valuation || buildValuation(metrics, marketData);
  const pattern = financials.cash_flow_pattern || metrics.cash_flow_pattern || classifyCashFlowPattern(metrics);
  const years = periods.map(period => period.year);
  sheet.columns = [
    { width: 5 }, { width: 14 }, { width: 16 }, { width: 22 }, { width: 16 }, { width: 16 }, { width: 16 },
    { width: 16 }, { width: 16 }, { width: 16 }, { width: 18 }, { width: 16 }, { width: 16 }, { width: 18 },
    { width: 34 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 34 }, { width: 14 },
    { width: 14 }, { width: 14 }, { width: 14 }, { width: 56 }, { width: 56 }, { width: 56 }
  ];
  sheet.getCell("B2").value = "Corporate analysis";
  sheet.getCell("B3").value = new Date().toISOString().slice(0, 10);
  sheet.getRow(4).values = [null, "Summary", null, "(단위:백만원)", ...years, "CAGR", "Financials", null, null, null, "기업 기본정보 / 사업 분석", "사업적 강점 / 리스크 요인", "Comments"];
  sheet.getCell("B5").value = "No.";
  sheet.getCell("C5").value = 1;
  sheet.getCell("K5").value = "(단위:백만원)";
  sheet.getCell("L5").value = "No. of Shares";
  sheet.getCell("M5").value = "Stock Price";
  sheet.getCell("N5").value = "Market Cap";
  sheet.getCell("O5").value = "현금흐름 분석";
  sheet.getCell("T5").value = "현금성자산 / 부채비율";
  sheet.getCell("Y5").value = `○ ${financials.company.name || financials.company.corp_code} 기업 개요와 사업 포지션`;
  sheet.getCell("Z5").value = "○ 사업적 강점";
  sheet.getCell("AA5").value = "DART API 기준 자동 생성. 주가/발행주식수는 별도 시장 데이터 기준일 확인 필요.";

  const rows = [
    [5, "매출액", 6],
    [6, "EBIT", 7],
    [7, "EBITDA", 10],
    [8, "당기순이익(손실)", 11],
    [12, "자산", 12],
    [13, "부채", 13],
    [14, "자본", 14],
    [15, "총차입금", 17],
    [16, "현금성자산", 18],
    [17, "순차입금(Net Debt)", 19],
    [19, "영업활동현금흐름", 20],
    [20, "투자활동현금흐름", 21],
    [21, "재무활동현금흐름", 22]
  ];
  for (const [row, label, sourceRow] of rows) {
    sheet.getCell(row, 4).value = label;
    periods.forEach((_, index) => {
      const col = 5 + index;
      const sourceCol = 2 + index;
      sheet.getCell(row, col).value = { formula: `'1'!${sheet.getCell(sourceRow, sourceCol).address}` };
    });
    if (periods.length >= 2) {
      const first = sheet.getCell(row, 5).address;
      const last = sheet.getCell(row, 4 + periods.length).address;
      sheet.getCell(row, 10).value = { formula: `IFERROR(RATE(${periods.length - 1},,-${first},${last}),"-")` };
    }
  }
  sheet.getCell("D9").value = "EBIT/매출액(%)";
  sheet.getCell("D10").value = "EBITDA Margin(%)";
  sheet.getCell("D11").value = "당기순이익(%)";
  for (const row of [9, 10, 11]) {
    periods.forEach((_, index) => {
      const col = 5 + index;
      const numerator = row === 9 ? 6 : row === 10 ? 7 : 8;
      sheet.getCell(row, col).value = { formula: `IFERROR(${sheet.getCell(numerator, col).address}/${sheet.getCell(5, col).address},"-")` };
    });
  }
  sheet.getCell("D18").value = "부채비율(%)";
  periods.forEach((_, index) => {
    const col = 5 + index;
    sheet.getCell(18, col).value = { formula: `IFERROR(${sheet.getCell(13, col).address}/${sheet.getCell(14, col).address},"-")` };
  });

  sheet.getCell("K6").value = "Market Cap";
  sheet.getCell("K7").value = "PER";
  sheet.getCell("K8").value = "EV";
  sheet.getCell("K9").value = "EV / EBITDA";
  sheet.getCell("L6").value = Number.isFinite(marketData.no_of_shares) ? marketData.no_of_shares : "리서치 필요";
  sheet.getCell("M6").value = Number.isFinite(marketData.stock_price) ? marketData.stock_price : "리서치 필요";
  sheet.getCell("N6").value = "발행주식수*주가";
  sheet.getCell("N7").value = "Market Cap / NI";
  sheet.getCell("N8").value = "Market Cap + Net Debt";
  sheet.getCell("N9").value = "EV / EBITDA";
  sheet.getCell("N6").value = { formula: `IFERROR(L6*M6/1000000,"입력 필요")` };
  sheet.getCell("N7").value = { formula: `IFERROR(N6/I8,"-")` };
  sheet.getCell("N8").value = { formula: `IFERROR(N6+I17,"-")` };
  sheet.getCell("N9").value = { formula: `IFERROR(N8/I7,"-")` };
  sheet.getCell("L6").fill = solid("FFF2CC");
  sheet.getCell("M6").fill = solid("FFF2CC");

  sheet.getCell("O17").value = `○영업현금흐름: ${formatWon(metrics.operating_cash_flow)}. ${pattern.type}: ${pattern.description}`;
  sheet.getCell("O19").value = `○투자현금흐름: ${formatWon(metrics.investing_cash_flow)}. CAPEX/M&A/금융상품 운용 여부를 분리 확인`;
  sheet.getCell("O21").value = `○재무현금흐름: ${formatWon(metrics.financing_cash_flow)}. 차입/상환/배당/증자 이벤트 확인`;
  sheet.getCell("Y14").value = "○ Industry Analysis";
  sheet.getCell("Y15").value = "DART 사업보고서의 사업부, 주요 제품, 고객, 시장 전망과 재무 변화 연결";
  sheet.getCell("Y16").value = "산업 성장률보다 매출 성장, 마진, 현금흐름으로 확인되는 회사 고유 수혜를 우선 해석";
  sheet.getCell("Z14").value = "○ 리스크 요인";
  sheet.getCell("Z15").value = "매출 둔화, 마진 하락, 재고/운전자본 부담, 차입금 증가, 투자회수 지연";
  sheet.getCell("AA14").value = "작성 가이드 반영";
  sheet.getCell("AA15").value = `DART 최신 사업보고서 우선, 5개년 확보 시 과거 공시 조회, 음수 부호 유지, D&A 주석 확인. ${formatMarketDataNote(marketData, valuation)}`;
  styleTemplateSheet(sheet);
}

function buildSummarySheet(sheet, financials) {
  const latest = [...financials.periods].reverse().find(period => period.metrics && Object.keys(period.metrics).length) || { metrics: {} };
  const metrics = latest.metrics;
  const pattern = financials.cash_flow_pattern || metrics.cash_flow_pattern || classifyCashFlowPattern(metrics);
  sheet.columns = [
    { width: 4 },
    { width: 18 },
    { width: 50 },
    { width: 4 },
    { width: 18 },
    { width: 26 }
  ];

  sheet.mergeCells("A1:F1");
  sheet.getCell("A1").value = `${financials.company.name || financials.company.corp_code} 주요 요약`;
  sheet.getCell("A1").font = { bold: true, size: 18, color: { argb: "FFFFFFFF" } };
  sheet.getCell("A1").fill = solid("17345F");
  sheet.getCell("A1").alignment = { vertical: "middle", horizontal: "left" };
  sheet.getRow(1).height = 34;

  sheet.getCell("B2").value = `${latest.year || financials.request.base_year}년`;
  sheet.getCell("E2").value = "매출";
  sheet.getCell("F2").value = formatWon(metrics.revenue);
  sheet.getCell("E3").value = "영업이익";
  sheet.getCell("F3").value = formatWon(metrics.operating_income);
  sheet.getCell("E4").value = "EBITDA";
  sheet.getCell("F4").value = formatWon(metrics.ebitda);
  sheet.getCell("E5").value = "순차입금";
  sheet.getCell("F5").value = formatWon(metrics.net_debt);
  sheet.getCell("E6").value = "현금흐름 패턴";
  sheet.getCell("F6").value = pattern.type;
  ["B2", "E2", "E3", "E4", "E5", "E6"].forEach(addr => sheet.getCell(addr).font = { bold: true });

  const sections = [
    ["C(고객)", [
      "> 주요 고객군과 매출 노출도를 사업보고서/IR에서 확인하세요.",
      "> 고객 집중도, 장기 공급계약, 고객사의 내재화 움직임을 리스크로 분리합니다."
    ]],
    ["P(상품)", [
      "> 주력 제품/서비스의 가격, 물량, 믹스 변화가 매출과 마진을 어떻게 움직이는지 봅니다.",
      "> 고부가 제품 전환 여부를 영업이익률 변화와 연결합니다."
    ]],
    ["C(채널)", [
      "> B2B 직판, 플랫폼, 대리점, JV 등 판매 채널이 현금흐름과 운전자본에 주는 영향을 정리합니다.",
      "> 해외 진출이나 현지 생산은 투자현금흐름과 함께 봅니다."
    ]],
    ["현금흐름 패턴", [
      `> ${pattern.type}: ${pattern.description}`,
      `> 영업 ${pattern.signs.operating || "?"} / 투자 ${pattern.signs.investing || "?"} / 재무 ${pattern.signs.financing || "?"}`
    ]]
  ];

  let row = 6;
  for (const [title, lines] of sections) {
    sheet.getCell(row, 2).value = title;
    sheet.getCell(row, 2).font = { bold: true, color: { argb: "FF17345F" } };
    row += 1;
    for (const line of lines) {
      sheet.mergeCells(row, 2, row, 6);
      sheet.getCell(row, 2).value = line;
      row += 1;
    }
    row += 1;
  }

  styleRange(sheet, "E2:F6");
}

function buildFinancialSheet(sheet, financials) {
  sheet.columns = [
    { header: "연도", key: "year", width: 12 },
    { header: "매출액(수익)", key: "revenue", width: 18 },
    { header: "영업이익(EBIT)", key: "operating_income", width: 18 },
    { header: "감가상각비", key: "depreciation", width: 18 },
    { header: "무형자산상각비", key: "amortization", width: 18 },
    { header: "EBITDA", key: "ebitda", width: 18 },
    { header: "영업이익률", key: "operating_margin", width: 14 },
    { header: "당기순이익", key: "net_income", width: 18 },
    { header: "자산총계", key: "total_assets", width: 18 },
    { header: "부채총계", key: "total_liabilities", width: 18 },
    { header: "자본총계", key: "total_equity", width: 18 },
    { header: "단기차입금", key: "short_term_debt", width: 18 },
    { header: "장기차입금", key: "long_term_debt", width: 18 },
    { header: "총차입금", key: "total_debt", width: 18 },
    { header: "현금성자산", key: "cash_and_equivalents", width: 18 },
    { header: "순차입금", key: "net_debt", width: 18 },
    { header: "부채비율", key: "debt_ratio", width: 14 },
    { header: "영업CF", key: "operating_cash_flow", width: 18 },
    { header: "투자CF", key: "investing_cash_flow", width: 18 },
    { header: "재무CF", key: "financing_cash_flow", width: 18 },
    { header: "현금흐름 패턴", key: "cash_flow_pattern", width: 18 }
  ];
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = solid("17345F");
  for (const period of financials.periods) {
    const m = period.metrics || {};
    sheet.addRow({
      year: period.year,
      revenue: m.revenue,
      operating_income: m.operating_income,
      depreciation: m.depreciation,
      amortization: m.amortization,
      ebitda: m.ebitda,
      operating_margin: m.operating_margin,
      net_income: m.net_income,
      total_assets: m.total_assets,
      total_liabilities: m.total_liabilities,
      total_equity: m.total_equity,
      short_term_debt: m.short_term_debt,
      long_term_debt: m.long_term_debt,
      total_debt: m.total_debt,
      cash_and_equivalents: m.cash_and_equivalents,
      net_debt: m.net_debt,
      debt_ratio: m.debt_ratio,
      operating_cash_flow: m.operating_cash_flow,
      investing_cash_flow: m.investing_cash_flow,
      financing_cash_flow: m.financing_cash_flow,
      cash_flow_pattern: (m.cash_flow_pattern && m.cash_flow_pattern.type) || ""
    });
  }
  ["B", "C", "D", "E", "F", "H", "I", "J", "K", "L", "M", "N", "O", "P", "R", "S", "T"].forEach(col => sheet.getColumn(col).numFmt = "#,##0");
  ["G", "Q"].forEach(col => sheet.getColumn(col).numFmt = "0.0%");
}

function buildMarketDataSheet(sheet, financials) {
  const marketData = financials.market_data || {};
  const valuation = financials.valuation || buildValuation((latestPeriod(financials).metrics || {}), marketData);
  sheet.columns = [{ width: 26 }, { width: 42 }, { width: 72 }];
  sheet.getRow(1).values = ["구분", "값", "출처/리서치 기준"];
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = solid("17345F");
  const sourceText = (marketData.sources || [])
    .map(source => `${source.type}: ${source.name}${source.symbol ? ` (${source.symbol})` : ""}${source.basis ? ` / ${source.basis}` : ""}`)
    .join("\n");
  sheet.addRows([
    ["No. of Shares", marketData.no_of_shares || "리서치 필요", "OpenDART stockTotqySttus의 보통주 발행주식총수 우선"],
    ["Stock Price", marketData.stock_price || "리서치 필요", "상장 시세 소스의 최신 거래일 종가/현재가"],
    ["Price Date", marketData.price_date || "확인 필요", "시세 기준일"],
    ["Market Cap", valuation.market_cap || "계산 대기", "No. of Shares x Stock Price"],
    ["PER", valuation.per || "계산 대기", "Market Cap / 최신 연도 당기순이익"],
    ["EV", valuation.enterprise_value || "계산 대기", "Market Cap + 최신 연도 순차입금"],
    ["EV/EBITDA", valuation.ev_ebitda || "계산 대기", "EV / 최신 연도 EBITDA"],
    ["자동 수집 출처", sourceText || "자동 수집 실패", "실패 시 Naver Finance/거래소/회사 IR 기준으로 AI 리서치 보완"],
    ["Warnings", (marketData.warnings || []).join("\n") || "없음", "리서치 보완이 필요한 셀을 Summary L6/M6에 유지"]
  ]);
  sheet.getColumn(2).numFmt = "#,##0.0";
  sheet.eachRow(row => {
    row.alignment = { vertical: "top", wrapText: true };
    row.eachCell(cell => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFE5E7EB" } },
        left: { style: "thin", color: { argb: "FFE5E7EB" } },
        bottom: { style: "thin", color: { argb: "FFE5E7EB" } },
        right: { style: "thin", color: { argb: "FFE5E7EB" } }
      };
    });
  });
}

function buildInsightsSheet(sheet, financials) {
  const latest = [...financials.periods].reverse().find(period => period.metrics && Object.keys(period.metrics).length);
  const insights = latest ? buildCompanyInsights(financials.company, latest, financials) : ["재무 데이터가 부족합니다."];
  sheet.columns = [{ width: 4 }, { width: 22 }, { width: 90 }];
  sheet.mergeCells("A1:C1");
  sheet.getCell("A1").value = "인사이트 + 산출물";
  sheet.getCell("A1").font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
  sheet.getCell("A1").fill = solid("17345F");
  const rows = [
    ["핵심 결론", insights[0] || ""],
    ["재무 안정성", insights[1] || ""],
    ["현금흐름 유형", insights[2] || ""],
    ["성장성 해석", insights[3] || ""],
    ["시장 데이터", insights[4] || ""],
    ["C/P/C 산출물", insights[5] || ""]
  ];
  rows.forEach((item, index) => {
    const row = index + 3;
    sheet.getCell(row, 2).value = item[0];
    sheet.getCell(row, 2).font = { bold: true, color: { argb: "FF17345F" } };
    sheet.getCell(row, 3).value = item[1];
  });
}

function buildLineageSheet(sheet, financials) {
  sheet.columns = [{ width: 24 }, { width: 80 }];
  const rows = [
    ["회사", financials.company.name || financials.company.corp_code],
    ["corp_code", financials.request.corp_code],
    ["stock_code", financials.company.stock_code || ""],
    ["no_of_shares", financials.market_data && financials.market_data.no_of_shares || ""],
    ["stock_price", financials.market_data && financials.market_data.stock_price || ""],
    ["price_date", financials.market_data && financials.market_data.price_date || ""],
    ["데이터 소스", financials.lineage.source],
    ["OpenDART endpoint", financials.lineage.endpoint],
    ["보고서 코드", financials.request.report_code],
    ["재무제표 기준", financials.request.fs_div],
    ["품질 점수", financials.quality.score],
    ["생성 시각", financials.lineage.generated_at]
  ];
  sheet.addRows(rows);
  sheet.getColumn(1).font = { bold: true };
}

function solid(argb) {
  return { type: "pattern", pattern: "solid", fgColor: { argb: `FF${argb}` } };
}

function styleRange(sheet, range) {
  const cells = sheet.getCell ? sheet.getRange : null;
  sheet.eachRow(row => {
    row.eachCell(cell => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFE5E7EB" } },
        bottom: { style: "thin", color: { argb: "FFE5E7EB" } }
      };
    });
  });
}

function styleTemplateSheet(sheet) {
  sheet.eachRow((row, rowNumber) => {
    row.alignment = { vertical: "middle", wrapText: true };
    row.eachCell(cell => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFE5E7EB" } },
        left: { style: "thin", color: { argb: "FFE5E7EB" } },
        bottom: { style: "thin", color: { argb: "FFE5E7EB" } },
        right: { style: "thin", color: { argb: "FFE5E7EB" } }
      };
      if (rowNumber <= 5 || ["K", "Y", "Z", "AA"].includes(cell.address.replace(/\d+/g, ""))) {
        cell.font = { ...(cell.font || {}), bold: true };
      }
    });
  });
  [2, 4, 5].forEach(row => {
    sheet.getRow(row).fill = solid("17345F");
    sheet.getRow(row).font = { bold: true, color: { argb: "FFFFFFFF" } };
  });
  sheet.getColumn("Y").alignment = { vertical: "top", wrapText: true };
  sheet.getColumn("Z").alignment = { vertical: "top", wrapText: true };
  sheet.getColumn("AA").alignment = { vertical: "top", wrapText: true };
}

function toMillion(value) {
  return Number.isFinite(value) ? Math.round(value / 1_000_000) : value;
}

function formatMultiple(value) {
  return Number.isFinite(value) ? `${value.toFixed(1)}x` : "확인 필요";
}

function formatMarketDataNote(marketData, valuation) {
  const shares = Number.isFinite(marketData.no_of_shares) ? `${marketData.no_of_shares.toLocaleString("ko-KR")}주` : "발행주식수 확인 필요";
  const price = Number.isFinite(marketData.stock_price) ? `${marketData.stock_price.toLocaleString("ko-KR")}원` : "주가 확인 필요";
  return `시장 데이터: ${shares}, ${price}, PER ${formatMultiple(valuation.per)}, EV/EBITDA ${formatMultiple(valuation.ev_ebitda)}.`;
}

function normalizeQuery(value) {
  return String(value || "")
    .toLowerCase()
    .replaceAll("주식회사", "")
    .replaceAll("(주)", "")
    .replaceAll("㈜", "")
    .replace(/[\s.,·_-]/g, "");
}

function safeJson(value, fallback) {
  try {
    return JSON.parse(value || "null") || fallback;
  } catch (error) {
    return fallback;
  }
}

function clamp(value, min, max) {
  if (!Number.isFinite(value)) return min;
  return Math.max(min, Math.min(max, Math.trunc(value)));
}

function formatWon(value) {
  if (!Number.isFinite(value)) return "확인 필요";
  if (Math.abs(value) >= 1_0000_0000_0000) return `${(value / 1_0000_0000_0000).toFixed(1)}조원`;
  if (Math.abs(value) >= 1_0000_0000) return `${(value / 1_0000_0000).toFixed(1)}억원`;
  return `${value.toLocaleString("ko-KR")}원`;
}

function formatPercent(value) {
  return Number.isFinite(value) ? `${(value * 100).toFixed(1)}%` : "확인 필요";
}
