const evidenceArchive = [
  {
    id: "E-001",
    title: "2024 Q4 Earnings Call",
    source: "IR Transcript",
    date: "2025-02-27",
    tag: "Revenue",
    reliability: 0.92,
    insight: "우주 시스템 부문 수주잔고 확대 언급"
  },
  {
    id: "E-002",
    title: "Electron Launch Manifest Update",
    source: "Company Blog",
    date: "2025-01-18",
    tag: "Launch Cadence",
    reliability: 0.73,
    insight: "연간 발사 빈도 목표 상향"
  },
  {
    id: "E-003",
    title: "Defense Contract News",
    source: "Industry News",
    date: "2025-02-09",
    tag: "Government",
    reliability: 0.65,
    insight: "국방계약 후보군 포함 보도"
  },
  {
    id: "E-004",
    title: "Neutron Infrastructure Filing",
    source: "Regulatory Filing",
    date: "2024-12-03",
    tag: "Neutron",
    reliability: 0.89,
    insight: "인프라 CAPEX 집행 일정 확인"
  }
];

const hypothesis = {
  title: "Neutron 상업화 가속 가능성",
  description:
    "발사 수요 증가 + 인프라 투자 집행 + 정부계약 모멘텀이 결합되면 12~18개월 내 매출 기여 시작 가능.",
  links: ["E-001", "E-003", "E-004"]
};

const yearlyFinancials = [
  { year: 2021, revenue: 62.2, cogsRatio: 83.5 },
  { year: 2022, revenue: 211.0, cogsRatio: 79.3 },
  { year: 2023, revenue: 244.6, cogsRatio: 74.8 },
  { year: 2024, revenue: 436.2, cogsRatio: 68.9 }
];

const financialSources = [
  "Rocket Lab USA FY2021~FY2024 Form 10-K",
  "Rocket Lab 공식 IR Earnings Presentation / Shareholder Letter"
];


const electronLaunchCadence = [
  { year: 2021, avgDays: 44.6 },
  { year: 2022, avgDays: 25.8 },
  { year: 2023, avgDays: 21.0 },
  { year: 2024, avgDays: 18.3 }
];

const stockTrend = [
  { month: "10월", close: 4.3 },
  { month: "11월", close: 4.8 },
  { month: "12월", close: 5.1 },
  { month: "1월", close: 5.4 },
  { month: "2월", close: 5.2 },
  { month: "3월", close: 5.6 }
];

const latestBacklog = {
  amount: "$1.07B",
  asOf: "2024 Q4",
  sourceLabel: "Rocket Lab earnings update"
};

const neutronRoadmap = [
  {
    phase: "개발 초기",
    milestone: "Archimedes 엔진 프로그램 구축 및 지상시험 착수",
    timing: "완료"
  },
  {
    phase: "인프라",
    milestone: "Wallops Launch Complex 3(LC-3) 및 생산 인프라 확장",
    timing: "진행 중"
  },
  {
    phase: "비행시험",
    milestone: "Neutron 첫 발사(official roadmap target)",
    timing: "타깃: 2025년 하반기"
  },
  {
    phase: "상업화",
    milestone: "정기 상업 임무 및 정부 임무로 전개",
    timing: "첫 발사 이후"
  }
];

const activeContracts = [
  {
    name: "미 국방부 소형 발사체 임무",
    type: "계약",
    stage: "기술 검토",
    amount: "$55M",
    due: "2025-06"
  },
  {
    name: "NASA 우주 시스템 공급",
    type: "입찰",
    stage: "RFP 제출 완료",
    amount: "$38M",
    due: "2025-05"
  },
  {
    name: "상업 위성 배치 패키지",
    type: "입찰",
    stage: "우선협상 대상",
    amount: "$27M",
    due: "2025-04"
  }
];

const acquisitions = [
  {
    company: "Sinclair Interplanetary",
    year: 2020,
    moat: "소형 위성용 반작용휠/스타트래커 수직계열화로 위성 플랫폼 통합 효율과 납기 통제력 강화"
  },
  {
    company: "Advanced Solutions Inc. (ASI)",
    year: 2021,
    moat: "비행소프트웨어(MAX Flight Software) 내재화로 우주선 소프트웨어 스택 잠금효과 및 전환비용 상승"
  },
  {
    company: "Solas Photon Counter (PSC)",
    year: 2021,
    moat: "우주급 태양전지 공급역량 확보로 핵심 전력부품 병목 완화 및 원가/성능 최적화"
  },
  {
    company: "Planetary Systems Corporation",
    year: 2021,
    moat: "분리시스템 포트폴리오 확보로 발사체-위성 인터페이스 표준 장악력 확대"
  }
];

const archiveList = document.querySelector("#archiveList");
const hypothesisTitle = document.querySelector("#hypothesisTitle");
const hypothesisDescription = document.querySelector("#hypothesisDescription");
const confidenceScore = document.querySelector("#confidenceScore");
const linkedEvidence = document.querySelector("#linkedEvidence");
const revenueLineChart = document.querySelector("#revenueLineChart");
const cogsLineChart = document.querySelector("#cogsLineChart");
const electronCadenceLineChart = document.querySelector("#electronCadenceLineChart");
const financialSource = document.querySelector("#financialSource");
const stockBars = document.querySelector("#stockBars");
const backlogChip = document.querySelector("#backlogChip");
const neutronRoadmapList = document.querySelector("#neutronRoadmapList");
const contractsList = document.querySelector("#contractsList");
const activeCount = document.querySelector("#activeCount");
const acquisitionsList = document.querySelector("#acquisitionsList");
const archiveSearch = document.querySelector("#archiveSearch");
const tagFilter = document.querySelector("#tagFilter");
const quickStats = document.querySelector("#quickStats");
const backToTop = document.querySelector("#backToTop");

let currentSort = "reliability";

function renderQuickStats() {
  const linkedItems = evidenceArchive.filter((item) => hypothesis.links.includes(item.id));
  const avgReliability =
    linkedItems.reduce((sum, item) => sum + item.reliability, 0) / Math.max(linkedItems.length, 1);

  const stats = [
    { label: "아카이브 증거", value: `${evidenceArchive.length}건` },
    { label: "최신 백로그", value: `${latestBacklog.amount}` },
    { label: "연간 매출(2024)", value: `$${yearlyFinancials.at(-1).revenue.toFixed(1)}M` },
    { label: "가설 신뢰도", value: `${(avgReliability * 100).toFixed(1)}%` }
  ];

  quickStats.innerHTML = "";
  stats.forEach((stat) => {
    const card = document.createElement("article");
    card.className = "stat-item";
    card.innerHTML = `
      <small>${stat.label}</small>
      <strong>${stat.value}</strong>
    `;
    quickStats.append(card);
  });
}

function renderArchive(items) {
  archiveList.innerHTML = "";

  if (items.length === 0) {
    archiveList.innerHTML = '<p class="empty">조건에 맞는 자료가 없습니다.</p>';
    return;
  }

  items.forEach((item) => {
    const card = document.createElement("article");
    card.className = "evidence-item";
    card.innerHTML = `
      <div class="evidence-top">
        <h3>${item.title}</h3>
        <span class="schedule-date">${item.date}</span>
      </div>
      <p>${item.insight}</p>
      <div class="meta">
        <span class="chip">${item.id}</span>
        <span class="chip">${item.source}</span>
        <span class="chip">${item.tag}</span>
        <span class="chip reliability">신뢰도 ${(item.reliability * 100).toFixed(0)}%</span>
      </div>
    `;

    archiveList.append(card);
  });
}

function calculateHypothesisScore() {
  const linkedItems = evidenceArchive.filter((item) => hypothesis.links.includes(item.id));
  const score =
    linkedItems.reduce((sum, item) => sum + item.reliability, 0) / Math.max(linkedItems.length, 1);
  return {
    linkedItems,
    score: (score * 100).toFixed(1)
  };
}

function renderHypothesis() {
  hypothesisTitle.textContent = hypothesis.title;
  hypothesisDescription.textContent = hypothesis.description;

  const { linkedItems, score } = calculateHypothesisScore();
  confidenceScore.textContent = `${score}%`;

  linkedEvidence.innerHTML = "";
  linkedItems.forEach((item) => {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = `${item.id} · ${item.title}`;
    linkedEvidence.append(chip);
  });
}

function renderLineChart(container, labels, values, color, formatter) {
  const width = 520;
  const height = 220;
  const padding = { top: 20, right: 20, bottom: 36, left: 42 };

  const minVal = Math.min(...values);
  const maxVal = Math.max(...values);
  const range = Math.max(maxVal - minVal, 1);

  const xStep = (width - padding.left - padding.right) / Math.max(values.length - 1, 1);

  const points = values.map((value, idx) => {
    const x = padding.left + idx * xStep;
    const y = height - padding.bottom - ((value - minVal) / range) * (height - padding.top - padding.bottom);
    return { x, y, value, label: labels[idx] };
  });

  const polylinePoints = points.map((point) => `${point.x},${point.y}`).join(" ");

  const xLabels = points
    .map(
      (point) =>
        `<text x="${point.x}" y="${height - 12}" text-anchor="middle" class="axis-label">${point.label}</text>`
    )
    .join("");

  const dots = points
    .map(
      (point) => `
        <circle cx="${point.x}" cy="${point.y}" r="4" fill="${color}"></circle>
        <text x="${point.x}" y="${point.y - 10}" text-anchor="middle" class="value-label">${formatter(point.value)}</text>
      `
    )
    .join("");

  container.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="line chart">
      <line x1="${padding.left}" y1="${height - padding.bottom}" x2="${width - padding.right}" y2="${height - padding.bottom}" class="axis"></line>
      <line x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${height - padding.bottom}" class="axis"></line>
      <polyline points="${polylinePoints}" class="series" style="stroke:${color}"></polyline>
      ${dots}
      ${xLabels}
    </svg>
  `;
}

function renderFinancials() {
  const labels = yearlyFinancials.map((row) => `${row.year}`);
  const revenueValues = yearlyFinancials.map((row) => row.revenue);
  const cogsValues = yearlyFinancials.map((row) => row.cogsRatio);
  const cadenceValues = electronLaunchCadence.map((row) => row.avgDays);
  const cadenceLabels = electronLaunchCadence.map((row) => `${row.year}`);

  renderLineChart(revenueLineChart, labels, revenueValues, "#1f73d4", (v) => `$${v.toFixed(1)}M`);
  renderLineChart(cogsLineChart, labels, cogsValues, "#0f9d58", (v) => `${v.toFixed(1)}%`);
  renderLineChart(electronCadenceLineChart, cadenceLabels, cadenceValues, "#d97706", (v) => `${v.toFixed(1)}d`);

  financialSource.textContent = `출처: ${financialSources.join(" / ")} / Electron cadence: Rocket Lab launch history(official mission log 기준)`;

  const maxPrice = Math.max(...stockTrend.map((entry) => entry.close));
  stockBars.innerHTML = "";

  stockTrend.forEach((entry) => {
    const bar = document.createElement("div");
    bar.className = "bar";
    bar.style.height = `${(entry.close / maxPrice) * 100}%`;
    bar.title = `${entry.month}: $${entry.close}`;
    bar.innerHTML = `
      <span>${entry.close}</span>
      <small>${entry.month}</small>
    `;
    stockBars.append(bar);
  });
}

function renderRoadmap() {
  backlogChip.textContent = `백로그 ${latestBacklog.amount} (${latestBacklog.asOf})`;
  neutronRoadmapList.innerHTML = "";

  neutronRoadmap.forEach((row) => {
    const item = document.createElement("article");
    item.className = "roadmap-item";
    item.innerHTML = `
      <div class="evidence-top">
        <h3>${row.phase}</h3>
        <span class="schedule-date">${row.timing}</span>
      </div>
      <p>${row.milestone}</p>
    `;
    neutronRoadmapList.append(item);
  });
}

function renderContracts() {
  contractsList.innerHTML = "";
  activeCount.textContent = `활성 ${activeContracts.length}건`;

  activeContracts.forEach((contract) => {
    const card = document.createElement("article");
    card.className = "contract-item";
    card.innerHTML = `
      <div class="evidence-top">
        <h3>${contract.name}</h3>
        <span class="schedule-date">${contract.due}</span>
      </div>
      <div class="meta">
        <span class="chip">유형: ${contract.type}</span>
        <span class="chip">단계: ${contract.stage}</span>
        <span class="chip">예상금액: ${contract.amount}</span>
      </div>
    `;
    contractsList.append(card);
  });
}

function renderAcquisitions() {
  acquisitionsList.innerHTML = "";

  acquisitions.forEach((item) => {
    const card = document.createElement("article");
    card.className = "acquisition-item";
    card.innerHTML = `
      <div class="evidence-top">
        <h3>${item.company}</h3>
        <span class="chip">${item.year}</span>
      </div>
      <p><strong>Moat:</strong> ${item.moat}</p>
    `;
    acquisitionsList.append(card);
  });
}

function byReliability(items) {
  return [...items].sort((a, b) => b.reliability - a.reliability);
}

function byRecent(items) {
  return [...items].sort((a, b) => new Date(b.date) - new Date(a.date));
}

function applyArchiveView() {
  const keyword = archiveSearch.value.trim().toLowerCase();
  const tag = tagFilter.value;

  const filtered = evidenceArchive.filter((item) => {
    const matchKeyword =
      item.title.toLowerCase().includes(keyword) || item.insight.toLowerCase().includes(keyword);
    const matchTag = tag === "all" ? true : item.tag === tag;
    return matchKeyword && matchTag;
  });

  const sorted = currentSort === "recent" ? byRecent(filtered) : byReliability(filtered);
  renderArchive(sorted);
}

function initTagFilter() {
  const tags = [...new Set(evidenceArchive.map((item) => item.tag))];
  tags.forEach((tag) => {
    const option = document.createElement("option");
    option.value = tag;
    option.textContent = tag;
    tagFilter.append(option);
  });
}

function setActiveSortButton(activeId) {
  document.querySelectorAll(".actions button").forEach((button) => {
    button.classList.toggle("active", button.id === activeId);
  });
}

document.querySelector("#sortReliability").addEventListener("click", () => {
  currentSort = "reliability";
  setActiveSortButton("sortReliability");
  applyArchiveView();
});

document.querySelector("#sortRecent").addEventListener("click", () => {
  currentSort = "recent";
  setActiveSortButton("sortRecent");
  applyArchiveView();
});

archiveSearch.addEventListener("input", applyArchiveView);
tagFilter.addEventListener("change", applyArchiveView);

window.addEventListener("scroll", () => {
  backToTop.classList.toggle("show", window.scrollY > 360);
});

backToTop.addEventListener("click", () => {
  window.scrollTo({ top: 0, behavior: "smooth" });
});

initTagFilter();
setActiveSortButton("sortReliability");
applyArchiveView();
renderQuickStats();
renderHypothesis();
renderFinancials();
renderRoadmap();
renderContracts();
renderAcquisitions();
