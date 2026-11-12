const state = {
  agency: "OPP",
  page: "dashboard",
  diagnostics: [],
  categories: [],
  skills: [],
  staff: [],
  programme: [],
  options: [],
  methodology: null,
};

const pageTitles = {
  dashboard: "Agency overview",
  roster: "Agency roster & participation",
  diagnostic: "CNA diagnostic",
  skills: "Skills gap review",
  needs: "Training needs analysis",
  programme: "Five-year training plan",
  talent: "Individual talent cards",
  monitoring: "Monitoring & evaluation",
  workspace: "Evidence workspace",
  method: "Analysis methodology",
};

const categories = [
  "Strategic Planning and Performance Management",
  "Leadership and Management Development",
  "Technical and Professional Development",
  "Governance, Ethics and Compliance",
  "Digital Skills and Innovation",
  "Career Development and Succession Planning",
  "GESI and Workplace Development",
];

const byId = (id) => document.getElementById(id);
const content = byId("pageContent");

function node(tag, className, value) {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (value !== undefined && value !== null) element.textContent = String(value);
  return element;
}

function setBanner(message, kind = "success") {
  const banner = byId("globalBanner");
  banner.hidden = !message;
  banner.className = `banner ${kind}`;
  banner.textContent = message || "";
}

function updateExportLinks() {
  byId("cnaReportLink").href = `/api/report.cna.pdf?agency=${encodeURIComponent(state.agency)}`;
  byId("excelExportLink").href = `/api/export.xlsx?agency=${encodeURIComponent(state.agency)}`;
}

async function api(path, options = {}) {
  const url = new URL(path, window.location.origin);
  if (url.pathname.startsWith("/api/") && !url.searchParams.has("agency")) {
    url.searchParams.set("agency", state.agency);
  }
  const response = await fetch(url, {
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    ...options,
  });
  const type = response.headers.get("content-type") || "";
  if (type.includes("application/json")) {
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || `Request failed (${response.status})`);
    return result;
  }
  if (!response.ok) throw new Error(`Request failed (${response.status})`);
  return response;
}

function formatNumber(value, decimals = 0) {
  if (value === null || value === undefined || value === "") return "—";
  return Number(value).toLocaleString(undefined, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

function sectionHeading(title, caption = "") {
  const wrapper = node("div", "section-heading");
  const textBox = node("div");
  textBox.append(node("h2", "", title));
  if (caption) textBox.append(node("p", "muted", caption));
  wrapper.append(textBox);
  return wrapper;
}

function card(title, value, note, accent = "") {
  const item = node("article", `metric-card ${accent}`);
  item.append(node("span", "metric-label", title));
  item.append(node("strong", "metric-value", value));
  if (note) item.append(node("span", "metric-note", note));
  return item;
}

function tableElement(headers, rows, emptyText = "No records found.") {
  const wrapper = node("div", "table-wrap");
  const table = node("table", "data-table");
  const thead = node("thead");
  const headerRow = node("tr");
  headers.forEach((value) => headerRow.append(node("th", "", value)));
  thead.append(headerRow);
  table.append(thead);
  const tbody = node("tbody");
  if (!rows.length) {
    const tr = node("tr");
    const td = node("td", "empty-cell", emptyText);
    td.colSpan = headers.length;
    tr.append(td);
    tbody.append(tr);
  } else {
    rows.forEach((row) => tbody.append(row));
  }
  table.append(tbody);
  wrapper.append(table);
  return wrapper;
}

function dataRow(values) {
  const row = node("tr");
  values.forEach((value) => {
    const cell = node("td");
    if (value instanceof Node) cell.append(value);
    else cell.textContent = value === null || value === undefined ? "—" : String(value);
    row.append(cell);
  });
  return row;
}

function progressBar(value, max = 10, color = "") {
  const wrap = node("div", "progress-wrap");
  const rail = node("div", "progress-rail");
  const fill = node("span", `progress-fill ${color}`);
  const percentage = max ? Math.max(0, Math.min(100, (Number(value) / max) * 100)) : 0;
  fill.style.width = `${percentage}%`;
  rail.append(fill);
  wrap.append(rail);
  return wrap;
}

function bandPill(value) {
  const band = node("span", `band-pill ${String(value || "").toLowerCase().replace(/[^a-z]+/g, "-")}`, value || "Not assessed");
  return band;
}

function statusPill(value) {
  const normalized = String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, "-");
  return node("span", `status-tag status-${normalized}`, value || "Draft");
}

async function loadPage(page = state.page) {
  state.page = page;
  byId("pageTitle").textContent = pageTitles[page] || "CNAMIS";
  document.querySelectorAll(".nav-link").forEach((button) => {
    button.classList.toggle("active", button.dataset.page === page);
  });
  byId("sidebar").classList.remove("sidebar-open");
  content.replaceChildren(node("div", "loading-card", "Loading local agency records…"));
  setBanner("");
  try {
    if (page === "dashboard") await renderDashboard();
    else if (page === "roster") await renderRoster();
    else if (page === "diagnostic") await renderDiagnostic();
    else if (page === "skills") await renderSkills();
    else if (page === "needs") await renderNeeds();
    else if (page === "programme") await renderProgramme();
    else if (page === "talent") await renderTalent();
    else if (page === "monitoring") await renderMonitoring();
    else if (page === "workspace") renderWorkspace();
    else if (page === "method") await renderMethodology();
  } catch (error) {
    content.replaceChildren(node("div", "error-card", error.message));
  }
}

async function renderDashboard() {
  const [stats, diagnostic, needs] = await Promise.all([
    api("/api/dashboard"),
    api("/api/diagnostic"),
    api("/api/needs"),
  ]);
  state.diagnostics = diagnostic;
  state.categories = needs;
  content.replaceChildren();

  const intro = node("div", "page-intro");
  intro.append(node("div", "eyebrow", "CAPABILITY NEEDS ANALYSIS · 2027–2031"));
  intro.append(node("h2", "", stats.agency.name));
  intro.append(node("p", "muted", "A source-grounded view of workforce coverage, CNA signals and the editable five-year development plan."));
  content.append(intro);

  const metrics = node("div", "metric-grid");
  metrics.append(
    card("Roster people", formatNumber(stats.roster_count), "Form 3.17 named occupants"),
    card("CNA submissions", formatNumber(stats.cna_responses), "Survey response records"),
    card("Roster matched", formatNumber(stats.matched_participants), `${stats.coverage_percent}% of named roster`, "metric-green"),
    card("No confirmed CNA match", formatNumber(stats.not_in_cna), "Follow up; name match is not proof of intent", "metric-orange"),
    card("Possible name matches", formatNumber(stats.verify_name), "Verify before classifying", "metric-blue"),
    card("SPA 4–5 responses", formatNumber(stats.spa_4_5_count), "Separate 1–5 scale; validate official SPA", "metric-purple"),
  );
  content.append(metrics);

  const grid = node("div", "two-column-grid");
  const assessmentCard = node("section", "panel");
  assessmentCard.append(sectionHeading("Lowest institutional CNA means", "Descriptive perception scores; lower means do not establish training as the remedy."));
  const lowRows = [...diagnostic].filter((item) => item.mean_score !== null).sort((a, b) => a.mean_score - b.mean_score).slice(0, 5);
  const bars = node("div", "bar-list");
  lowRows.forEach((item) => {
    const line = node("div", "bar-line");
    const copy = node("div", "bar-copy");
    copy.append(node("strong", "", `${item.item_code} · ${item.statement}`));
    copy.append(node("span", "muted", `Gap ${formatNumber(item.gap, 2)} · ${item.band}`));
    line.append(copy, progressBar(item.mean_score, 10, "fill-blue"), node("b", "bar-value", formatNumber(item.mean_score, 2)));
    bars.append(line);
  });
  assessmentCard.append(bars);

  const needsCard = node("section", "panel");
  needsCard.append(sectionHeading("Most-mentioned development areas", "Phrase-matched CNA E7 interests; categories overlap."));
  const topNeeds = [...needs].sort((a, b) => b.respondents_mentioning - a.respondents_mentioning).slice(0, 5);
  const needBars = node("div", "bar-list");
  topNeeds.forEach((item) => {
    const line = node("div", "bar-line");
    line.append(node("span", "bar-name", item.name));
    line.append(progressBar(item.respondents_mentioning, stats.cna_responses, "fill-green"));
    line.append(node("b", "bar-value", `${item.respondents_mentioning}/${stats.cna_responses}`));
    needBars.append(line);
  });
  needsCard.append(needBars);
  grid.append(assessmentCard, needsCard);
  content.append(grid);

  const notice = node("div", "notice notice-info");
  notice.append(node("strong", "", "Planning assumptions"));
  notice.append(node("p", "", "The 2025–2030 corporate plan is image-based in the supplied files; 2031 successor priorities, budgets, provider quotes and eligibility rules remain unconfirmed. Edit and validate the draft before approval."));
  content.append(notice);
}

async function renderDiagnostic() {
  const items = await api("/api/diagnostic");
  state.diagnostics = items;
  content.replaceChildren();
  content.append(sectionHeading("Institutional assessment by survey item", "CNA score bands: 1–3 Poor/Low · 4–8 Average · 9–10 High/Excellent. Capability gap = 10 − mean score."));
  const dimensionStats = new Map();
  items.forEach((item) => {
    const current = dimensionStats.get(item.dimension) || { sum: 0, count: 0 };
    if (item.mean_score !== null) {
      current.sum += Number(item.mean_score);
      current.count += 1;
    }
    dimensionStats.set(item.dimension, current);
  });
  const dimensions = node("div", "dimension-grid");
  dimensionStats.forEach((value, label) => {
    const mean = value.count ? value.sum / value.count : null;
    const block = node("article", "dimension-card");
    block.append(node("span", "metric-label", label));
    block.append(node("strong", "metric-value", mean === null ? "—" : formatNumber(mean, 2)));
    block.append(node("span", "metric-note", mean === null ? "Not assessed" : `${scoreBand(mean)} · gap ${formatNumber(10 - mean, 2)}`));
    dimensions.append(block);
  });
  content.append(dimensions);
  const rows = items.map((item) => dataRow([
    item.dimension,
    item.item_code,
    item.statement,
    formatNumber(item.valid_scores),
    item.mean_score === null ? "Not assessed" : formatNumber(item.mean_score, 2),
    item.gap === null ? "—" : formatNumber(item.gap, 2),
    bandPill(item.band),
    item.review_route,
  ]));
  content.append(tableElement(["Dimension", "Item", "Survey statement", "n", "Mean / 10", "Gap", "Band", "Review route"], rows));
  const note = node("div", "notice notice-warning");
  note.append(node("strong", "", "Interpretation"));
  note.append(node("p", "", "Scores reflect respondent views of agency statements. Confirm causes and decide whether the response is training, policy reform, management action or institutional strengthening."));
  content.append(note);
}

function scoreBand(score) {
  return score <= 3 ? "Poor/Low" : score <= 8 ? "Average" : "High/Excellent";
}

async function renderSkills() {
  const records = await api("/api/skills");
  state.skills = records;
  content.replaceChildren();
  content.append(sectionHeading("Qualifications, reported competencies and SPA", "Qualifications are compared manually with the verbatim job requirement. No equivalency, job competency standard or expected performance target was supplied."));
  const controls = node("div", "filter-row");
  const search = node("input", "input-control");
  search.placeholder = "Filter by officer, position or division…";
  search.setAttribute("aria-label", "Filter skills records");
  const note = node("span", "muted", `${records.length} CNA respondent records · SPA remains on its separate 1–5 scale`);
  controls.append(search, note);
  const tableSlot = node("div");
  content.append(controls, tableSlot);
  const draw = () => {
    const query = search.value.toLowerCase();
    const visible = records.filter((row) => `${row.name} ${row.position} ${row.division}`.toLowerCase().includes(query));
    const rows = visible.map((item) => dataRow([
      item.name, item.position, item.division, item.qualification, item.required_qualification,
      item.qualification_review, item.competencies, item.spa_rating === null ? "Not recorded" : `${item.spa_rating} / 5`,
    ]));
    tableSlot.replaceChildren(tableElement(["Respondent", "Position", "Division", "Qualification held", "Position requirement", "Review", "Reported competencies", "SPA"], rows));
  };
  search.addEventListener("input", draw);
  draw();
}

async function renderNeeds() {
  const [needs, dash] = await Promise.all([api("/api/needs"), api("/api/dashboard")]);
  state.categories = needs;
  content.replaceChildren();
  content.append(sectionHeading("Training need signals", "Counts come from phrase matching against the supplied E7 topic-interest answers. They are overlapping and not formal priorities."));
  const note = node("div", "notice notice-info");
  note.append(node("strong", "", "Priority decisions are left for agency review."));
  note.append(node("p", "", "Use severity, corporate-plan importance, operational criticality, number of staff affected and urgency. The supplied records do not support a defensible weighted priority score."));
  content.append(note);
  const cards = node("div", "need-grid");
  needs.forEach((item) => {
    const tile = node("article", "need-card");
    tile.append(node("span", "need-count", `${item.respondents_mentioning} / ${dash.cna_responses}`));
    tile.append(node("h3", "", item.name));
    tile.append(node("p", "muted", item.intervention_route));
    const terms = node("div", "keyword-list");
    item.keywords.forEach((word) => terms.append(node("span", "keyword", word)));
    tile.append(terms);
    tile.append(node("small", "muted", "Source: CNA E7 · phrase rules shown; categories may overlap."));
    cards.append(tile);
  });
  content.append(cards);
  const evidence = node("div", "panel panel-compact");
  evidence.append(node("h3", "", "Needs identification process"));
  evidence.append(node("p", "", `${dash.needs_process_yes} respondents answered Yes and ${dash.needs_process_no} answered No to whether the agency has a process for identifying future training needs.`));
  content.append(evidence);
}

async function renderProgramme() {
  const [programme, options, needs] = await Promise.all([api("/api/programme"), api("/api/options"), api("/api/needs")]);
  state.programme = programme;
  state.options = options;
  state.categories = needs;
  content.replaceChildren();
  content.append(sectionHeading("Draft agency programme · 2027–2031", "Edit interventions, timing, estimated PGK budget and approval state. Amounts remain unpriced until an approved provider quote is entered."));

  const model = node("div", "learning-model");
  model.append(
    learningPillar("10%", "Formal", "Conditional classroom/provider learning after a verified gap, eligibility and approval."),
    learningPillar("20%", "Social", "Coaching, mentoring and peer learning tied to workplace needs."),
    learningPillar("70%", "Experiential", "Supervised practice and applied work-based learning."),
  );
  content.append(model);

  const scenarioPanel = node("section", "panel");
  scenarioPanel.append(sectionHeading("Existing plan options", "Alternative planning approaches retained from the workbook; these are not approved plans."));
  const scenarioGrid = node("div", "scenario-grid");
  options.forEach((option) => {
    const details = node("details", "scenario-card");
    const summary = node("summary", "");
    summary.append(node("strong", "", option.title));
    summary.append(node("span", "muted", option.description));
    details.append(summary);
    const schedule = node("div", "scenario-years");
    option.annual_schedule.forEach((year) => {
      const yearBlock = node("div", "scenario-year");
      yearBlock.append(node("b", "", year.year), node("span", "", year.focus));
      schedule.append(yearBlock);
    });
    details.append(schedule);
    scenarioGrid.append(details);
  });
  scenarioPanel.append(scenarioGrid);
  content.append(scenarioPanel);

  const filters = node("div", "filter-row");
  const yearSelect = node("select", "input-control");
  yearSelect.append(optionElement("", "All plan years"));
  [2027, 2028, 2029, 2030, 2031].forEach((year) => yearSelect.append(optionElement(String(year), String(year))));
  const categorySelect = node("select", "input-control");
  categorySelect.append(optionElement("", "All categories"));
  categories.forEach((item) => categorySelect.append(optionElement(item.name, item.name)));
  filters.append(yearSelect, categorySelect, node("span", "muted", `${programme.length} editable activities`));
  const tableSlot = node("div");
  content.append(filters, tableSlot);
  const draw = () => {
    const visible = programme.filter((item) =>
      (!yearSelect.value || Number(item.year) === Number(yearSelect.value))
      && (!categorySelect.value || item.category === categorySelect.value)
    );
    const rows = visible.map((item) => {
      const edit = node("button", "button button-small button-outline", "Edit");
      edit.type = "button";
      edit.addEventListener("click", () => openPlanEditor(item));
      return dataRow([
        item.year, item.category, item.focus, item.owner_role, item.timing,
        item.estimated_budget === null ? "Not priced" : `PGK ${formatNumber(item.estimated_budget, 2)}`,
        statusPill(item.approval_status), edit,
      ]);
    });
    tableSlot.replaceChildren(tableElement(["Year", "Category", "Draft focus", "Proposed owner", "Timing", "Estimate", "Approval", "Action"], rows));
  };
  yearSelect.addEventListener("change", draw);
  categorySelect.addEventListener("change", draw);
  draw();
  const reportLink = node("a", "button button-primary", "Download Capability Needs Analysis Report (PDF)");
  reportLink.href = `/api/report.cna.pdf?agency=${encodeURIComponent(state.agency)}`;
  reportLink.download = "OPP-Capability-Needs-Analysis-Report.pdf";
  const exportLink = node("a", "button button-outline", "Download current Excel report");
  exportLink.href = `/api/export.xlsx?agency=${encodeURIComponent(state.agency)}`;
  exportLink.download = "CNAMIS-OPP-report.xlsx";
  const downloadRow = node("div", "download-row");
  downloadRow.append(reportLink, exportLink);
  content.append(downloadRow);
}

async function renderRoster() {
  const [roster, pending] = await Promise.all([api("/api/staff"), api("/api/matches")]);
  state.staff = roster;
  content.replaceChildren();
  content.append(sectionHeading("Roster participation reconciliation", "Form 3.17 named occupants reconciled with CNA names. Verify possible one-name matches before changing their status."));
  const counts = {
    participated: roster.filter((item) => item.cna_status === "Participated").length,
    review: roster.filter((item) => item.cna_status === "Verify name").length,
    missing: roster.filter((item) => item.cna_status === "Not in CNA").length,
  };
  const metrics = node("div", "metric-grid metric-grid-three");
  metrics.append(
    card("Participated", counts.participated, "Confirmed name match", "metric-green"),
    card("Verify name", counts.review, "Human identity check required", "metric-orange"),
    card("No CNA match", counts.missing, "No confirmed matching survey name", "metric-blue"),
  );
  content.append(metrics);
  const filter = node("div", "filter-row");
  const search = node("input", "input-control");
  search.placeholder = "Filter by officer, position or location…";
  const status = node("select", "input-control");
  status.append(optionElement("", "All statuses"), optionElement("Participated", "Participated"), optionElement("Verify name", "Verify name"), optionElement("Not in CNA", "Not in CNA"));
  filter.append(search, status);
  const tableSlot = node("div");
  content.append(filter, tableSlot);
  const draw = () => {
    const query = search.value.toLowerCase();
    const visible = roster.filter((item) =>
      (!status.value || item.cna_status === status.value)
      && `${item.name} ${item.positions.join(" ")} ${item.locations.join(" ")}`.toLowerCase().includes(query)
    );
    const rows = visible.map((item) => dataRow([
      item.name, item.position_numbers.join(", "), item.positions.join("; "),
      item.locations.join("; "), statusPill(item.cna_status),
    ]));
    tableSlot.replaceChildren(tableElement(["Roster name", "Position number", "Position / designation", "Location", "CNA status"], rows));
  };
  search.addEventListener("input", draw);
  status.addEventListener("change", draw);
  draw();

  const reviewPanel = node("section", "panel");
  reviewPanel.append(sectionHeading("Possible matches to verify", "A name-only response is not confirmed as roster participation until it is manually checked against agency records."));
  if (!pending.length) {
    reviewPanel.append(node("p", "muted", "There are no pending name matches."));
  } else {
    pending.forEach((item) => {
      const line = node("div", "review-line");
      const description = node("div");
      description.append(node("strong", "", `${item.respondent} · survey row ${item.source_row}`));
      description.append(node("small", "muted", "Confirm against identity and agency records; do not rely on a first name alone."));
      const select = node("select", "input-control");
      select.append(optionElement("", "Choose roster officer…"));
      item.candidates.forEach((candidate) => {
        const label = `${candidate.name} · ${candidate.positions.join("; ")}${candidate.locations.length ? ` · ${candidate.locations.join(", ")}` : ""}`;
        select.append(optionElement(String(candidate.id), label));
      });
      const verify = node("button", "button button-small button-primary", "Verify match");
      verify.type = "button";
      verify.disabled = item.candidates.length === 0;
      verify.addEventListener("click", async () => {
        if (!select.value) {
          setBanner("Choose the matching roster officer first.", "error");
          return;
        }
        try {
          const result = await api("/api/matches/verify", {
            method: "POST",
            body: JSON.stringify({ agency: state.agency, response_id: item.response_id, staff_id: Number(select.value) }),
          });
          setBanner(result.message);
          await renderRoster();
        } catch (error) {
          setBanner(error.message, "error");
        }
      });
      line.append(description, select, verify);
      reviewPanel.append(line);
    });
  }
  content.append(reviewPanel);
}

function learningPillar(percent, label, description) {
  const pillar = node("article", "learning-pillar");
  pillar.append(node("strong", "", percent), node("b", "", label), node("span", "muted", description));
  return pillar;
}

function optionElement(value, label) {
  const option = node("option", "", label);
  option.value = value;
  return option;
}

function openPlanEditor(activity) {
  const dialog = byId("editDialog");
  byId("editActivityId").textContent = `${activity.activity_id} · ${activity.year} · ${activity.category}`;
  const host = byId("editFields");
  host.replaceChildren();
  const fields = [
    { key: "year", label: "Year", type: "select", options: ["2027", "2028", "2029", "2030", "2031"] },
    { key: "category", label: "Training category", type: "select", options: categories },
    { key: "focus", label: "Annual focus / intervention", type: "textarea" },
    { key: "formal_learning", label: "10% formal learning (conditional)", type: "textarea" },
    { key: "social_learning", label: "20% social learning", type: "textarea" },
    { key: "experiential_learning", label: "70% experiential learning", type: "textarea" },
    { key: "owner_role", label: "Proposed owner role", type: "text" },
    { key: "timing", label: "Timing / quarter", type: "text" },
    { key: "estimated_budget", label: "Estimated budget (PGK)", type: "number", min: "0", step: "0.01" },
    { key: "quote_basis", label: "Quote / cost basis", type: "text" },
    { key: "corporate_reference", label: "Approved corporate-plan / KRA reference", type: "text" },
    { key: "approval_status", label: "Approval status", type: "select", options: ["Draft - review and approval required", "Submitted for review", "Approved", "In progress", "Completed", "On hold"] },
    { key: "indicator", label: "Indicator / target (confirm)", type: "textarea" },
    { key: "outcome_notes", label: "Outcome / evidence notes", type: "textarea" },
  ];
  fields.forEach((field) => {
    const label = node("label", "form-field");
    label.append(node("span", "", field.label));
    let input;
    if (field.type === "select") {
      input = node("select", "input-control");
      field.options.forEach((value) => input.append(optionElement(value, value)));
      input.value = String(activity[field.key] ?? "");
    } else if (field.type === "textarea") {
      input = node("textarea", "input-control");
      input.value = activity[field.key] ?? "";
      input.rows = 3;
    } else {
      input = node("input", "input-control");
      input.type = field.type;
      input.value = activity[field.key] ?? "";
      if (field.min) input.min = field.min;
      if (field.step) input.step = field.step;
    }
    input.name = field.key;
    input.dataset.field = field.key;
    label.append(input);
    host.append(label);
  });
  byId("savePlanButton").onclick = async () => {
    const changes = {};
    host.querySelectorAll("[data-field]").forEach((input) => {
      changes[input.dataset.field] = input.value;
    });
    changes.year = Number(changes.year);
    changes.estimated_budget = changes.estimated_budget === "" ? null : Number(changes.estimated_budget);
    try {
      await api(`/api/programme/${encodeURIComponent(activity.activity_id)}`, {
        method: "PATCH",
        body: JSON.stringify({ agency: state.agency, ...changes }),
      });
      dialog.close();
      setBanner("Plan activity updated.");
      await renderProgramme();
    } catch (error) {
      setBanner(error.message, "error");
    }
  };
  dialog.showModal();
}

async function renderTalent() {
  const records = await api("/api/skills");
  state.skills = records;
  content.replaceChildren();
  content.append(sectionHeading("CNA respondents", "Cards show supplied qualifications, role, SPA, career goals and interests; eligibility remains subject to agency rules."));
  const search = node("input", "input-control");
  search.placeholder = "Search officer, position or division…";
  const count = node("span", "muted", `${records.length} respondents`);
  const controls = node("div", "filter-row");
  controls.append(search, count);
  const tableSlot = node("div");
  content.append(controls, tableSlot);
  const draw = () => {
    const q = search.value.toLowerCase();
    const visible = records.filter((row) => `${row.name} ${row.position} ${row.division}`.toLowerCase().includes(q));
    const rows = visible.map((item) => {
      const button = node("button", "button button-small button-outline", "View card");
      button.type = "button";
      button.addEventListener("click", () => openTalent(item));
      return dataRow([item.name, item.position, item.division, item.employment_status, item.spa_rating === null ? "Not recorded" : `${item.spa_rating} / 5`, button]);
    });
    tableSlot.replaceChildren(tableElement(["Officer", "Position", "Division", "Employment status", "SPA 1–5", "Profile"], rows));
  };
  search.addEventListener("input", draw);
  draw();
}

async function openTalent(item) {
  try {
    const profile = await api(`/api/respondents/${item.response_id}`);
    byId("talentName").textContent = profile.name;
    const host = byId("talentDetails");
    host.replaceChildren();
    const details = [
      ["Position", profile.position],
      ["Grade", profile.grade],
      ["Division / section", profile.division],
      ["Employment status", profile.employment_status || "Not recorded"],
      ["Highest qualification", profile.qualification || "Not recorded"],
      ["Position qualification requirement", profile.required_qualification || "Not recorded"],
      ["Qualification review", profile.qualification_comparison],
      ["SPA rating", profile.spa_rating === null ? "Not recorded (scale 1–5)" : `${profile.spa_rating} / 5`],
      ["Training eligibility", profile.training_eligibility],
      ["Competencies reported", profile.competencies || "Not recorded"],
      ["Career goal", profile.career_goal || "Not recorded"],
      ["Selected interests", profile.interests || "Not recorded"],
      ["CNA name match", profile.match_status],
      ["Source row", profile.source_row],
    ];
    details.forEach(([label, value]) => {
      const block = node("div", "profile-item");
      block.append(node("span", "muted", label), node("strong", "", value));
      host.append(block);
    });
    byId("talentDialog").showModal();
  } catch (error) {
    setBanner(error.message, "error");
  }
}

async function renderMonitoring() {
  const programme = await api("/api/programme");
  state.programme = programme;
  content.replaceChildren();
  content.append(sectionHeading("Implementation and outcomes", "Update participants, expenditure, indicators, targets and evidence as records become available."));
  const summary = node("div", "notice notice-warning");
  summary.append(node("strong", "", "No training delivery or spend is assumed."));
  summary.append(node("p", "", "All monitoring actuals are blank until entered by the agency. Budget variance is calculated only after both an estimate and actual spend are recorded."));
  content.append(summary);
  const rows = programme.map((item) => {
    const completed = item.participants_completed;
    const planned = item.planned_participants;
    const rate = planned && completed !== null ? `${formatNumber((completed / planned) * 100, 1)}%` : "—";
    const edit = node("button", "button button-small button-outline", "Update");
    edit.type = "button";
    edit.addEventListener("click", () => openMonitorEditor(item));
    return dataRow([
      item.year, item.activity_id, item.category, item.indicator || "Not set",
      planned ?? "—", completed ?? "—", rate,
      item.actual_expenditure === null ? "—" : `PGK ${formatNumber(item.actual_expenditure, 2)}`,
      item.actual_result || "Not reported", statusPill(item.status), edit,
    ]);
  });
  content.append(tableElement(["Year", "Activity", "Category", "Indicator", "Planned", "Completed", "Rate", "Actual spend", "Outcome", "Status", "Action"], rows));
}

function openMonitorEditor(activity) {
  const dialog = byId("monitorDialog");
  byId("monitorActivityId").textContent = `${activity.activity_id} · ${activity.year} · ${activity.category}`;
  const host = byId("monitorFields");
  host.replaceChildren();
  const fields = [
    { key: "planned_participants", label: "Planned participants", type: "number", min: "0" },
    { key: "participants_completed", label: "Participants completed", type: "number", min: "0" },
    { key: "actual_expenditure", label: "Actual expenditure (PGK)", type: "number", min: "0", step: "0.01" },
    { key: "monitor_indicator", label: "Indicator", type: "text" },
    { key: "baseline", label: "Baseline", type: "text" },
    { key: "target", label: "Target", type: "text" },
    { key: "actual_result", label: "Actual result", type: "text" },
    { key: "evidence", label: "Evidence / corrective action", type: "textarea" },
  ];
  fields.forEach((field) => {
    const label = node("label", "form-field");
    label.append(node("span", "", field.label));
    const input = field.type === "textarea" ? node("textarea", "input-control") : node("input", "input-control");
    input.name = field.key;
    input.dataset.field = field.key;
    if (field.type !== "textarea") {
      input.type = field.type;
      if (field.min) input.min = field.min;
      if (field.step) input.step = field.step;
    } else {
      input.rows = 3;
    }
    const currentKey = field.key === "monitor_indicator" ? "monitor_indicator" : field.key;
    input.value = activity[currentKey] ?? "";
    label.append(input);
    host.append(label);
  });
  byId("saveMonitorButton").onclick = async () => {
    const data = { activity_id: activity.activity_id, agency: state.agency };
    host.querySelectorAll("[data-field]").forEach((input) => {
      const key = input.dataset.field;
      const number = input.type === "number";
      data[key] = number && input.value !== "" ? Number(input.value) : input.value;
    });
    if (Object.hasOwn(data, "monitor_indicator")) {
      data.indicator = data.monitor_indicator;
      delete data.monitor_indicator;
    }
    try {
      await api("/api/monitoring", { method: "POST", body: JSON.stringify(data) });
      dialog.close();
      setBanner("Monitoring record saved.");
      await renderMonitoring();
    } catch (error) {
      setBanner(error.message, "error");
    }
  };
  dialog.showModal();
}

function renderWorkspace() {
  content.replaceChildren();
  content.append(sectionHeading("Ask questions about local OPP records", "Responses are generated by local evidence queries. No external AI service is connected."));
  const panel = node("section", "workspace-panel");
  const prompt = node("p", "muted", "Try: “What are the lowest CNA scores?”, “How many officers did not match?”, “What are the leadership needs?”, or ask about a named respondent.");
  const form = node("form", "workspace-form");
  const input = node("textarea", "input-control");
  input.placeholder = "Ask about agency participation, CNA gaps, training needs, an officer, or budget status…";
  input.rows = 3;
  input.required = true;
  const submit = node("button", "button button-primary", "Ask local records");
  submit.type = "submit";
  form.append(input, submit);
  const answerBox = node("div", "answer-list");
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!input.value.trim()) return;
    submit.disabled = true;
    submit.textContent = "Searching…";
    try {
      const answer = await api("/api/workspace", {
        method: "POST",
        body: JSON.stringify({ agency: state.agency, question: input.value }),
      });
      const answerCard = node("article", "answer-card");
      answerCard.append(node("p", "", answer.answer));
      answerCard.append(node("small", "muted", `Sources: ${answer.sources.join(" · ")}`));
      answerBox.prepend(answerCard);
      input.value = "";
    } catch (error) {
      setBanner(error.message, "error");
    } finally {
      submit.disabled = false;
      submit.textContent = "Ask local records";
    }
  });
  panel.append(prompt, form, answerBox);
  content.append(panel);
}

async function renderMethodology() {
  const data = await api("/api/methodology");
  state.methodology = data;
  content.replaceChildren();
  content.append(sectionHeading("How the training plan was developed", "Transparent, source-bound steps used by the local CNAMIS application."));
  const steps = [
    ["01", "Load source records", data.sources.join(" ")],
    ["02", "Reconcile CNA participation", "Match Form 3.17 occupants to CNA names using family and given-name tokens. Possible one-name matches are held for verification; missing matches are flagged, not treated as proof of non-response."],
    ["03", "Diagnose institutional ratings", `${data.cna_scale} ${data.gap_formula}`],
    ["04", "Review individual skills evidence", `${data.qualification_method} ${data.competency_method}`],
    ["05", "Classify expressed needs", "Phrase-match E7 topic interests to the seven requested training categories. Counts overlap and are not treated as confirmed gaps or staff affected."],
    ["06", "Draft the five-year programme", "The 2027–2031 activities come from the supplied CNAMIS workbook. Each combines formal, social and work-based learning approaches. Users can edit the focus, year, timing, owner, budget basis, approval and indicators."],
    ["07", "Prioritise with agency validation", data.priority_method],
    ["08", "Monitor implementation", "The M&E tracker accepts participant, expenditure, baseline, target, result and evidence updates. It does not populate outcomes until an agency user supplies them."],
  ];
  const timeline = node("div", "method-list");
  steps.forEach(([number, heading, body]) => {
    const item = node("article", "method-step");
    item.append(node("span", "method-number", number));
    const textBox = node("div");
    textBox.append(node("h3", "", heading), node("p", "muted", body));
    item.append(textBox);
    timeline.append(item);
  });
  content.append(timeline);
  const facts = [
    ["Learning model", data.learning_model],
    ["Training eligibility", data.training_eligibility],
    ["Budget", data.budget_method],
    ["Corporate-plan limitation", data.corporate_plan_limit],
    ["Local workspace", data.local_only],
    ["Evidence workspace", data.workspace_mode],
  ];
  const factGrid = node("div", "fact-grid");
  facts.forEach(([label, value]) => {
    const item = node("article", "fact-card");
    item.append(node("b", "", label), node("p", "muted", value));
    factGrid.append(item);
  });
  content.append(factGrid);
}

async function init() {
  try {
    const agencies = await api("/api/agencies");
    const select = byId("agencySelect");
    select.replaceChildren();
    agencies.forEach((agency) => {
      const option = optionElement(agency.code, agency.name);
      select.append(option);
    });
    if (!agencies.some((agency) => agency.code === state.agency)) {
      state.agency = agencies[0]?.code || "OPP";
    }
    select.value = state.agency;
    updateExportLinks();
    select.addEventListener("change", () => {
      state.agency = select.value;
      updateExportLinks();
      loadPage();
    });
    document.querySelectorAll(".nav-link, .brand").forEach((button) => {
      button.addEventListener("click", (event) => {
        event.preventDefault();
        loadPage(button.dataset.page || "dashboard");
      });
    });
    byId("menuButton").addEventListener("click", () => byId("sidebar").classList.toggle("sidebar-open"));
    await loadPage("dashboard");
  } catch (error) {
    content.replaceChildren(node("div", "error-card", `CNAMIS could not load local data: ${error.message}`));
  }
}

init();
