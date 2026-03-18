const STORAGE_KEY = "attendance-performance-system/v3";
const TIME_SLOTS = ["09:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00", "18:00"];
const SUPABASE_TABLE = "app_states";
const SAVE_DEBOUNCE_MS = 300;

const els = {
  dateInput: document.getElementById("dateInput"),
  teamFilter: document.getElementById("teamFilter"),
  quickSearch: document.getElementById("quickSearch"),
  teamForm: document.getElementById("teamForm"),
  teamEditingId: document.getElementById("teamEditingId"),
  teamName: document.getElementById("teamName"),
  teamLeader: document.getElementById("teamLeader"),
  teamDailyDeviceTarget: document.getElementById("teamDailyDeviceTarget"),
  teamMonthlyDeviceTarget: document.getElementById("teamMonthlyDeviceTarget"),
  teamSubmitBtn: document.getElementById("teamSubmitBtn"),
  teamTableBody: document.getElementById("teamTableBody"),
  employeeForm: document.getElementById("employeeForm"),
  employeeEditingId: document.getElementById("employeeEditingId"),
  employeeName: document.getElementById("employeeName"),
  employeeTeam: document.getElementById("employeeTeam"),
  employeeMonthlyTarget: document.getElementById("employeeMonthlyTarget"),
  employeeSubmitBtn: document.getElementById("employeeSubmitBtn"),
  employeeTableBody: document.getElementById("employeeTableBody"),
  attendanceHead: document.getElementById("attendanceHead"),
  attendanceBody: document.getElementById("attendanceBody"),
  performanceBody: document.getElementById("performanceBody"),
  summaryBody: document.getElementById("summaryBody"),
  teamBoard: document.getElementById("teamBoard"),
  slotLegend: document.getElementById("slotLegend"),
  dailyAttendanceTotal: document.getElementById("dailyAttendanceTotal"),
  dailyAchievedTotal: document.getElementById("dailyAchievedTotal"),
  monthlyAchievedTotal: document.getElementById("monthlyAchievedTotal"),
  monthlyCompletionRate: document.getElementById("monthlyCompletionRate"),
  saveStatus: document.getElementById("saveStatus"),
  cloudBadge: document.getElementById("cloudBadge"),
  authSummary: document.getElementById("authSummary"),
  authStatus: document.getElementById("authStatus"),
  authEmail: document.getElementById("authEmail"),
  authPassword: document.getElementById("authPassword"),
  signInBtn: document.getElementById("signInBtn"),
  signUpBtn: document.getElementById("signUpBtn"),
  signOutBtn: document.getElementById("signOutBtn"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportBtn: document.getElementById("exportBtn"),
  importBtn: document.getElementById("importBtn"),
  importInput: document.getElementById("importInput"),
  resetBtn: document.getElementById("resetBtn"),
};

let state = loadState();
let lastSavedAt = null;
let saveTimer = null;
let saveQueue = Promise.resolve();

const cloud = {
  client: null,
  configured: false,
  session: null,
  channel: null,
  authSubscription: null,
  hydrated: false,
  signature: "",
};

function formatDate(date) {
  return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
}

function getToday() {
  return formatDate(new Date());
}

function createId() {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID();
  }

  return `id-${Date.now()}-${Math.random().toString(16).slice(2, 10)}`;
}

function getSeedData() {
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const todayKey = formatDate(today);
  const yesterdayKey = formatDate(yesterday);
  const defaultTeam = "梅州-霍白超团队";

  return {
    teams: [
      {
        id: createId(),
        name: defaultTeam,
        leader: "霍白超",
        dailyDeviceTarget: 26,
        monthlyDeviceTarget: 335,
      },
    ],
    employees: [
      { id: createId(), name: "丁宁", team: defaultTeam, dailyTarget: 5, monthlyTarget: 80 },
      { id: createId(), name: "古雁和", team: defaultTeam, dailyTarget: 5, monthlyTarget: 60 },
      { id: createId(), name: "江伟业", team: defaultTeam, dailyTarget: 4, monthlyTarget: 45 },
      { id: createId(), name: "张超", team: defaultTeam, dailyTarget: 4, monthlyTarget: 60 },
      { id: createId(), name: "邹春生", team: defaultTeam, dailyTarget: 3, monthlyTarget: 45 },
      { id: createId(), name: "郑文峰", team: defaultTeam, dailyTarget: 5, monthlyTarget: 45 },
      { id: createId(), name: "谢烈雄", team: defaultTeam, dailyTarget: 3, monthlyTarget: 45 },
      { id: createId(), name: "黄文舒", team: defaultTeam, dailyTarget: 4, monthlyTarget: 50 },
      { id: createId(), name: "江友帮", team: defaultTeam, dailyTarget: 4, monthlyTarget: 50 },
      { id: createId(), name: "阅瑞梅", team: defaultTeam, dailyTarget: 3, monthlyTarget: 40 },
      { id: createId(), name: "杨梓豪", team: defaultTeam, dailyTarget: 3, monthlyTarget: 40 },
      { id: createId(), name: "曾浩", team: defaultTeam, dailyTarget: 3, monthlyTarget: 40 },
    ],
    attendance: {},
    performance: {},
  };
}

function buildDerivedTeams(employees, existingTeams = []) {
  const teamsByName = new Map();

  for (const employee of employees) {
    if (!teamsByName.has(employee.team)) {
      teamsByName.set(employee.team, {
        id: createId(),
        name: employee.team,
        leader: "",
        dailyDeviceTarget: 0,
        monthlyDeviceTarget: 0,
      });
    }
  }

  for (const team of existingTeams) {
    teamsByName.set(team.name, {
      id: team.id || createId(),
      name: team.name,
      leader: team.leader || "",
      dailyDeviceTarget: Number(team.dailyDeviceTarget || 0),
      monthlyDeviceTarget: Number(team.monthlyDeviceTarget || 0),
    });
  }

  return [...teamsByName.values()].sort((a, b) => a.name.localeCompare(b.name, "zh-CN"));
}

function normalizeState(rawState) {
  const normalized = {
    teams: Array.isArray(rawState?.teams) ? rawState.teams : [],
    employees: Array.isArray(rawState?.employees) ? rawState.employees : [],
    attendance: rawState?.attendance ?? {},
    performance: rawState?.performance ?? {},
  };

  normalized.teams = buildDerivedTeams(normalized.employees, normalized.teams);
  return normalized;
}

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return getSeedData();
  }

  try {
    return normalizeState(JSON.parse(raw));
  } catch (error) {
    console.warn("读取本地数据失败，已恢复示例数据", error);
    return getSeedData();
  }
}

function getStateSignature(value = state) {
  return JSON.stringify(normalizeState(value));
}

function formatSaveTime(date) {
  return date.toLocaleTimeString("zh-CN", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

function updateSaveStatus(kind) {
  const hasCloudSession = Boolean(cloud.session);
  let text = "自动保存已开启";

  if (kind === "saving") {
    text = hasCloudSession ? "正在保存到云端..." : "正在保存到本地...";
  } else if (kind === "saved") {
    const suffix = lastSavedAt ? ` ${formatSaveTime(lastSavedAt)}` : "";
    if (hasCloudSession) {
      text = `已同步到云端${suffix}`;
    } else if (cloud.configured) {
      text = `已本地保存${suffix}，登录后可同步云端`;
    } else {
      text = `已本地保存${suffix}`;
    }
  } else if (kind === "error") {
    text = hasCloudSession ? "云端保存失败，已保留本地副本" : "本地保存失败，请尽快导出备份";
  } else if (hasCloudSession) {
    text = "云端同步已开启";
  } else if (cloud.configured) {
    text = "自动保存已开启，登录后可同步云端";
  }

  els.saveStatus.dataset.state = kind;
  els.saveStatus.textContent = text;
}

function setAuthStatus(message, tone = "") {
  els.authStatus.textContent = message;
  els.authStatus.className = tone ? `auth-status ${tone}` : "auth-status";
}

function setCloudBadge(text, tone) {
  els.cloudBadge.textContent = text;
  els.cloudBadge.className = `cloud-badge ${tone}`;
}

function persistLocalState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(normalizeState(state)));
  lastSavedAt = new Date();
}

async function persistCloudState(force = false) {
  if (!cloud.client || !cloud.session || !cloud.hydrated) {
    return;
  }

  const payload = normalizeState(state);
  const signature = getStateSignature(payload);
  if (!force && signature === cloud.signature) {
    return;
  }

  const { data, error } = await cloud.client
    .from(SUPABASE_TABLE)
    .upsert(
      {
        user_id: cloud.session.user.id,
        payload,
      },
      { onConflict: "user_id" },
    )
    .select("updated_at");

  if (error) {
    throw error;
  }

  cloud.signature = signature;
  const updatedAt = Array.isArray(data) ? data[0]?.updated_at : data?.updated_at;
  lastSavedAt = updatedAt ? new Date(updatedAt) : new Date();
}

async function persistState(forceCloud = false) {
  try {
    persistLocalState();
    await persistCloudState(forceCloud);
    updateSaveStatus("saved");
  } catch (error) {
    console.error("保存数据失败", error);
    updateSaveStatus("error");
  }
}

function queuePersist(forceCloud = false) {
  saveQueue = saveQueue
    .catch(() => undefined)
    .then(() => persistState(forceCloud));

  return saveQueue;
}

function scheduleSave({ render = true, immediate = false, forceCloud = false } = {}) {
  if (render) {
    renderAll();
  }

  updateSaveStatus("saving");
  if (saveTimer !== null) {
    window.clearTimeout(saveTimer);
  }

  if (immediate) {
    saveTimer = null;
    void queuePersist(forceCloud);
    return;
  }

  saveTimer = window.setTimeout(() => {
    saveTimer = null;
    void queuePersist(forceCloud);
  }, SAVE_DEBOUNCE_MS);
}

function flushPendingSave() {
  if (saveTimer !== null) {
    window.clearTimeout(saveTimer);
    saveTimer = null;
    void queuePersist();
  }
}

function uniqueTeams() {
  return [...new Set([...state.teams.map((team) => team.name), ...state.employees.map((employee) => employee.team)])].sort((a, b) =>
    a.localeCompare(b, "zh-CN"),
  );
}

function getTeamProfile(teamName) {
  return state.teams.find((team) => team.name === teamName) ?? {
    id: createId(),
    name: teamName,
    leader: "",
    dailyDeviceTarget: 0,
    monthlyDeviceTarget: 0,
  };
}

function getEmployeeProfile(employeeName) {
  return state.employees.find((employee) => employee.name === employeeName) ?? null;
}

function ensureDailyAttendance(dateKey, employeeName) {
  state.attendance[dateKey] ??= {};
  state.attendance[dateKey][employeeName] ??= {};

  for (const slot of TIME_SLOTS) {
    state.attendance[dateKey][employeeName][slot] ??= false;
  }

  return state.attendance[dateKey][employeeName];
}

function ensureDailyPerformance(dateKey, employeeName) {
  state.performance[dateKey] ??= {};
  state.performance[dateKey][employeeName] ??= {
    assigned: 0,
    selfDeveloped: 0,
    dailyTarget: Number(getEmployeeProfile(employeeName)?.dailyTarget || 0),
  };
  state.performance[dateKey][employeeName].assigned ??= 0;
  state.performance[dateKey][employeeName].selfDeveloped ??= 0;
  state.performance[dateKey][employeeName].dailyTarget ??= Number(getEmployeeProfile(employeeName)?.dailyTarget || 0);
  return state.performance[dateKey][employeeName];
}

function getFilteredEmployees() {
  const team = els.teamFilter.value;
  const keyword = els.quickSearch.value.trim();

  return state.employees
    .filter((employee) => !team || employee.team === team)
    .filter((employee) => !keyword || employee.name.includes(keyword))
    .sort((a, b) => a.name.localeCompare(b.name, "zh-CN"));
}

function groupEmployeesByTeam(employees) {
  const groups = new Map();

  for (const employee of employees) {
    if (!groups.has(employee.team)) {
      groups.set(employee.team, []);
    }
    groups.get(employee.team).push(employee);
  }

  return [...groups.entries()]
    .sort((a, b) => a[0].localeCompare(b[0], "zh-CN"))
    .map(([team, members]) => ({
      team,
      members: members.sort((a, b) => a.name.localeCompare(b.name, "zh-CN")),
    }));
}

function monthPrefix(dateKey) {
  return dateKey.slice(0, 7);
}

function calculateAttendanceCount(dateKey, employeeName) {
  const record = ensureDailyAttendance(dateKey, employeeName);
  return TIME_SLOTS.reduce((total, slot) => total + (record[slot] ? 1 : 0), 0);
}

function calculateMonthlyPerformance(dateKey, employeeName) {
  const prefix = monthPrefix(dateKey);
  let assigned = 0;
  let selfDeveloped = 0;

  for (const [recordDate, employees] of Object.entries(state.performance)) {
    if (!recordDate.startsWith(prefix) || !employees[employeeName]) {
      continue;
    }

    assigned += Number(employees[employeeName].assigned || 0);
    selfDeveloped += Number(employees[employeeName].selfDeveloped || 0);
  }

  return {
    assigned,
    selfDeveloped,
    achieved: assigned + selfDeveloped,
  };
}

function getCompletionClass(ratio) {
  if (ratio >= 1) {
    return "ratio-good";
  }
  if (ratio >= 0.6) {
    return "ratio-normal";
  }
  return "ratio-low";
}

function getStatusLabel(attendanceCount, dailyRatio, monthlyRatio) {
  if (dailyRatio >= 1 && monthlyRatio >= 0.8) {
    return { text: "表现稳定", className: "status-excellent" };
  }
  if (attendanceCount >= 4 || dailyRatio >= 0.6 || monthlyRatio >= 0.5) {
    return { text: "持续跟进", className: "status-progress" };
  }
  return { text: "需要关注", className: "status-risk" };
}

function percent(value) {
  return `${(Number(value || 0) * 100).toFixed(2)}%`;
}

function getEmployeeMetrics(dateKey, employee) {
  const attendanceCount = calculateAttendanceCount(dateKey, employee.name);
  const daily = ensureDailyPerformance(dateKey, employee.name);
  const todayAmount = Number(daily.assigned || 0) + Number(daily.selfDeveloped || 0);
  const dailyTarget = Number(daily.dailyTarget || 0);
  const dailyRatio = dailyTarget ? todayAmount / dailyTarget : 0;
  const monthly = calculateMonthlyPerformance(dateKey, employee.name);
  const monthlyRatio = employee.monthlyTarget ? monthly.achieved / employee.monthlyTarget : 0;

  return {
    attendanceCount,
    daily,
    dailyTarget,
    todayAmount,
    dailyRatio,
    monthly,
    monthlyRatio,
    status: getStatusLabel(attendanceCount, dailyRatio, monthlyRatio),
  };
}

function getTeamMetrics(dateKey, employees) {
  const totals = employees.reduce(
    (current, employee) => {
      const metrics = getEmployeeMetrics(dateKey, employee);
      current.headcount += 1;
      current.attendance += metrics.attendanceCount;
      current.todayAmount += metrics.todayAmount;
      current.monthlyAchieved += metrics.monthly.achieved;
      return current;
    },
    { headcount: 0, attendance: 0, todayAmount: 0, monthlyAchieved: 0 },
  );

  const profile = employees[0] ? getTeamProfile(employees[0].team) : null;
  const dailyDeviceTarget = Number(profile?.dailyDeviceTarget || 0);
  const monthlyDeviceTarget = Number(profile?.monthlyDeviceTarget || 0);

  return {
    ...totals,
    teamProfile: profile,
    dailyDeviceTarget,
    monthlyDeviceTarget,
    dailyDeviceCompletion: dailyDeviceTarget ? totals.todayAmount / dailyDeviceTarget : 0,
    monthlyDeviceCompletion: monthlyDeviceTarget ? totals.monthlyAchieved / monthlyDeviceTarget : 0,
  };
}

function renderTeamFilter() {
  const selected = els.teamFilter.value;
  const teams = uniqueTeams();
  els.teamFilter.innerHTML = ['<option value="">全部团队</option>', ...teams.map((team) => `<option value="${team}">${team}</option>`)].join("");
  els.teamFilter.value = teams.includes(selected) || selected === "" ? selected : "";
}

function renderEmployeeTeamOptions() {
  const selected = els.employeeTeam.value;
  const teams = uniqueTeams();
  els.employeeTeam.innerHTML = ['<option value="">请选择团队</option>', ...teams.map((team) => `<option value="${team}">${team}</option>`)].join("");
  els.employeeTeam.value = teams.includes(selected) ? selected : "";
}

function renderLegend() {
  els.slotLegend.innerHTML = TIME_SLOTS.map((slot) => `<span class="chip">${slot}</span>`).join("");
}

function renderAuthPanel() {
  if (!cloud.configured) {
    setCloudBadge("未配置", "cloud-warning");
    els.authSummary.textContent = "还没有配置 Supabase 项目，当前页面只会做本地保存。";
    setAuthStatus("请先填写 supabase.config.js 或在 GitHub Actions 中配置 SUPABASE_URL / SUPABASE_ANON_KEY。", "warning");
    els.signOutBtn.classList.add("hidden");
    els.signInBtn.classList.remove("hidden");
    els.signUpBtn.classList.remove("hidden");
    return;
  }

  if (cloud.session?.user) {
    setCloudBadge("云端在线", "cloud-online");
    els.authSummary.textContent = `当前账号：${cloud.session.user.email}，已开启云端实时同步。`;
    setAuthStatus("账号状态正常。当前浏览器修改的数据会自动同步到 Supabase。", "success");
    els.signOutBtn.classList.remove("hidden");
    els.signInBtn.classList.add("hidden");
    els.signUpBtn.classList.add("hidden");
    return;
  }

  setCloudBadge("等待登录", "cloud-idle");
  els.authSummary.textContent = "Supabase 已配置完成。登录后会自动把本地数据同步到云端。";
  setAuthStatus("如果项目开启了邮箱确认，注册后请先去邮箱完成验证。", "warning");
  els.signOutBtn.classList.add("hidden");
  els.signInBtn.classList.remove("hidden");
  els.signUpBtn.classList.remove("hidden");
}

function renderTeams() {
  if (!state.teams.length) {
    els.teamTableBody.innerHTML = '<tr class="empty-row"><td colspan="5">当前还没有团队档案。</td></tr>';
    return;
  }

  els.teamTableBody.innerHTML = [...state.teams]
    .sort((a, b) => a.name.localeCompare(b.name, "zh-CN"))
    .map(
      (team) => `
        <tr>
          <td>${team.name}</td>
          <td>${team.leader || "-"}</td>
          <td>${team.dailyDeviceTarget}</td>
          <td>${team.monthlyDeviceTarget}</td>
          <td>
            <button class="tiny-button secondary" data-action="edit-team" data-id="${team.id}">编辑</button>
            <button class="tiny-button danger" data-action="delete-team" data-id="${team.id}">删除</button>
          </td>
        </tr>
      `,
    )
    .join("");
}

function renderEmployees() {
  const teamGroups = groupEmployeesByTeam(getFilteredEmployees());

  if (!teamGroups.length) {
    els.employeeTableBody.innerHTML = '<tr class="empty-row"><td colspan="4">当前筛选条件下没有员工。</td></tr>';
    return;
  }

  els.employeeTableBody.innerHTML = teamGroups
    .flatMap(({ team, members }) => [
      `<tr class="group-row"><td colspan="4">${team} · 共 ${members.length} 人 · 团队长 ${getTeamProfile(team).leader || "-"}</td></tr>`,
      ...members.map(
        (employee) => `
          <tr>
            <td>${employee.name}</td>
            <td>${employee.team}</td>
            <td>${employee.monthlyTarget}</td>
            <td>
              <button class="tiny-button secondary" data-action="edit-employee" data-id="${employee.id}">编辑</button>
              <button class="tiny-button danger" data-action="delete-employee" data-id="${employee.id}">删除</button>
            </td>
          </tr>
        `,
      ),
    ])
    .join("");
}

function renderAttendance() {
  const dateKey = els.dateInput.value;
  const teamGroups = groupEmployeesByTeam(getFilteredEmployees());

  els.attendanceHead.innerHTML = `
    <tr>
      <th>序号</th>
      <th>团队归属</th>
      <th>作业BD</th>
      ${TIME_SLOTS.map((slot) => `<th>${slot}</th>`).join("")}
      <th>出勤次数</th>
    </tr>
  `;

  if (!teamGroups.length) {
    els.attendanceBody.innerHTML = `<tr class="empty-row"><td colspan="${TIME_SLOTS.length + 4}">没有可显示的出勤数据。</td></tr>`;
    return;
  }

  let index = 0;
  els.attendanceBody.innerHTML = teamGroups
    .flatMap(({ team, members }) => {
      const teamTotals = getTeamMetrics(dateKey, members);
      const subtotalRow = `<tr class="group-row"><td colspan="${TIME_SLOTS.length + 4}">${team} · 团队长 ${teamTotals.teamProfile?.leader || "-"} · ${members.length} 人 · 团队出勤合计 ${teamTotals.attendance}</td></tr>`;
      const memberRows = members.map((employee) => {
        index += 1;
        const record = ensureDailyAttendance(dateKey, employee.name);
        const attendanceCount = calculateAttendanceCount(dateKey, employee.name);

        return `
          <tr>
            <td>${index}</td>
            <td>${employee.team}</td>
            <td>${employee.name}</td>
            ${TIME_SLOTS.map((slot) => `
              <td class="checkbox-cell">
                <input
                  type="checkbox"
                  data-action="toggle-attendance"
                  data-name="${employee.name}"
                  data-slot="${slot}"
                  ${record[slot] ? "checked" : ""}
                />
              </td>
            `).join("")}
            <td>${attendanceCount}</td>
          </tr>
        `;
      });

      return [subtotalRow, ...memberRows];
    })
    .join("");
}

function renderPerformance() {
  const dateKey = els.dateInput.value;
  const teamGroups = groupEmployeesByTeam(getFilteredEmployees());

  if (!teamGroups.length) {
    els.performanceBody.innerHTML = '<tr class="empty-row"><td colspan="11">没有可显示的业绩数据。</td></tr>';
    return;
  }

  els.performanceBody.innerHTML = teamGroups
    .flatMap(({ team, members }) => {
      const teamTotals = getTeamMetrics(dateKey, members);
      const subtotalRow = `<tr class="group-row"><td colspan="11">${team} · 团队长 ${teamTotals.teamProfile?.leader || "-"} · 日设备目标 ${teamTotals.dailyDeviceTarget} · 今日量 ${teamTotals.todayAmount} · 日完成率 ${percent(teamTotals.dailyDeviceCompletion)} · 月设备目标 ${teamTotals.monthlyDeviceTarget} · 月完成率 ${percent(teamTotals.monthlyDeviceCompletion)}</td></tr>`;

      const memberRows = members.map((employee) => {
        const { daily, dailyTarget, todayAmount, dailyRatio, monthly, monthlyRatio } = getEmployeeMetrics(dateKey, employee);

        return `
          <tr>
            <td>${employee.name}</td>
            <td><input type="number" min="0" step="1" data-action="performance-input" data-field="assigned" data-name="${employee.name}" value="${daily.assigned}" /></td>
            <td><input type="number" min="0" step="1" data-action="performance-input" data-field="selfDeveloped" data-name="${employee.name}" value="${daily.selfDeveloped}" /></td>
            <td>${todayAmount}</td>
            <td><input type="number" min="0" step="1" data-action="performance-input" data-field="dailyTarget" data-name="${employee.name}" value="${dailyTarget}" /></td>
            <td class="${getCompletionClass(dailyRatio)}">${percent(dailyRatio)}</td>
            <td>${monthly.assigned}</td>
            <td>${monthly.selfDeveloped}</td>
            <td>${monthly.achieved}</td>
            <td>${employee.monthlyTarget}</td>
            <td class="${getCompletionClass(monthlyRatio)}">${percent(monthlyRatio)}</td>
          </tr>
        `;
      });

      return [subtotalRow, ...memberRows];
    })
    .join("");
}

function renderSummary() {
  const dateKey = els.dateInput.value;
  const teamGroups = groupEmployeesByTeam(getFilteredEmployees());

  if (!teamGroups.length) {
    els.summaryBody.innerHTML = '<tr class="empty-row"><td colspan="11">没有可显示的整合数据。</td></tr>';
    return;
  }

  let totalAttendance = 0;
  let totalAchieved = 0;
  let totalMonthlyAchieved = 0;
  let totalMonthlyTarget = 0;

  const rows = teamGroups.flatMap(({ team, members }) => {
    const teamTotals = getTeamMetrics(dateKey, members);
    totalAttendance += teamTotals.attendance;
    totalAchieved += teamTotals.todayAmount;
    totalMonthlyAchieved += teamTotals.monthlyAchieved;
    totalMonthlyTarget += teamTotals.monthlyDeviceTarget;

    const subtotalRow = `
      <tr class="group-row">
        <td>${dateKey}</td>
        <td>${team}</td>
        <td>团队小计 / ${teamTotals.teamProfile?.leader || "-"}</td>
        <td>${teamTotals.attendance}</td>
        <td>${teamTotals.todayAmount}</td>
        <td>${teamTotals.dailyDeviceTarget}</td>
        <td>${percent(teamTotals.dailyDeviceCompletion)}</td>
        <td>${teamTotals.monthlyAchieved}</td>
        <td>${teamTotals.monthlyDeviceTarget}</td>
        <td>${percent(teamTotals.monthlyDeviceCompletion)}</td>
        <td>团队汇总</td>
      </tr>
    `;

    const memberRows = members.map((employee) => {
      const metrics = getEmployeeMetrics(dateKey, employee);

      return `
        <tr>
          <td>${dateKey}</td>
          <td>${employee.team}</td>
          <td>${employee.name}</td>
          <td>${metrics.attendanceCount}</td>
          <td>${metrics.todayAmount}</td>
          <td>${metrics.dailyTarget}</td>
          <td class="${getCompletionClass(metrics.dailyRatio)}">${percent(metrics.dailyRatio)}</td>
          <td>${metrics.monthly.achieved}</td>
          <td>${employee.monthlyTarget}</td>
          <td class="${getCompletionClass(metrics.monthlyRatio)}">${percent(metrics.monthlyRatio)}</td>
          <td><span class="status-pill ${metrics.status.className}">${metrics.status.text}</span></td>
        </tr>
      `;
    });

    return [subtotalRow, ...memberRows];
  });

  els.summaryBody.innerHTML = rows.join("");
  els.dailyAttendanceTotal.textContent = String(totalAttendance);
  els.dailyAchievedTotal.textContent = String(totalAchieved);
  els.monthlyAchievedTotal.textContent = String(totalMonthlyAchieved);
  els.monthlyCompletionRate.textContent = totalMonthlyTarget ? percent(totalMonthlyAchieved / totalMonthlyTarget) : "0%";
}

function renderTeamBoard() {
  const dateKey = els.dateInput.value;
  const teamGroups = groupEmployeesByTeam(getFilteredEmployees());

  if (!teamGroups.length) {
    els.teamBoard.innerHTML = '<article class="team-card"><p>当前筛选条件下没有团队数据。</p></article>';
    return;
  }

  els.teamBoard.innerHTML = teamGroups
    .map(({ team, members }) => {
      const totals = getTeamMetrics(dateKey, members);

      return `
        <article class="team-card">
          <h3>${team}</h3>
          <p>团队长：${totals.teamProfile?.leader || "-"} ｜ 成员：${members.map((member) => member.name).join("、")}</p>
          <div class="team-metrics">
            <div class="team-metric"><span>团队人数</span><strong>${totals.headcount}</strong></div>
            <div class="team-metric"><span>当日出勤</span><strong>${totals.attendance}</strong></div>
            <div class="team-metric"><span>今日完成量</span><strong>${totals.todayAmount}</strong></div>
            <div class="team-metric"><span>日设备目标</span><strong>${totals.dailyDeviceTarget}</strong></div>
            <div class="team-metric"><span>日完成率</span><strong>${percent(totals.dailyDeviceCompletion)}</strong></div>
            <div class="team-metric"><span>月设备目标</span><strong>${totals.monthlyDeviceTarget}</strong></div>
            <div class="team-metric"><span>月度达成</span><strong>${totals.monthlyAchieved}</strong></div>
            <div class="team-metric"><span>月完成率</span><strong>${percent(totals.monthlyDeviceCompletion)}</strong></div>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderAll() {
  renderAuthPanel();
  renderTeamFilter();
  renderEmployeeTeamOptions();
  renderLegend();
  renderTeamBoard();
  renderTeams();
  renderEmployees();
  renderAttendance();
  renderPerformance();
  renderSummary();
}

function syncAfterChange() {
  scheduleSave({ render: true, immediate: true });
}

function resetTeamForm() {
  els.teamForm.reset();
  els.teamEditingId.value = "";
  els.teamSubmitBtn.textContent = "新增团队";
}

function resetEmployeeForm() {
  els.employeeForm.reset();
  els.employeeEditingId.value = "";
  els.employeeSubmitBtn.textContent = "新增员工";
}

function addTeam(event) {
  event.preventDefault();
  const editingId = els.teamEditingId.value;
  const name = els.teamName.value.trim();
  const leader = els.teamLeader.value.trim();
  const dailyDeviceTarget = Number(els.teamDailyDeviceTarget.value);
  const monthlyDeviceTarget = Number(els.teamMonthlyDeviceTarget.value);

  if (!name || !leader) {
    return;
  }

  const duplicate = state.teams.find((team) => team.name === name && team.id !== editingId);
  if (duplicate) {
    alert("该团队名称已存在，请换一个名称或直接编辑原团队。");
    return;
  }

  if (editingId) {
    const previousName = state.teams.find((team) => team.id === editingId)?.name;
    state.teams = state.teams.map((team) =>
      team.id === editingId
        ? {
            ...team,
            name,
            leader,
            dailyDeviceTarget,
            monthlyDeviceTarget,
          }
        : team,
    );

    if (previousName && previousName !== name) {
      state.employees = state.employees.map((employee) =>
        employee.team === previousName
          ? {
              ...employee,
              team: name,
            }
          : employee,
      );
    }
  } else {
    state.teams.push({
      id: createId(),
      name,
      leader,
      dailyDeviceTarget,
      monthlyDeviceTarget,
    });
  }

  resetTeamForm();
  syncAfterChange();
}

function addEmployee(event) {
  event.preventDefault();
  const editingId = els.employeeEditingId.value;
  const name = els.employeeName.value.trim();
  const team = els.employeeTeam.value.trim();
  const monthlyTarget = Number(els.employeeMonthlyTarget.value);

  if (!name || !team) {
    return;
  }

  const duplicate = state.employees.find((employee) => employee.name === name && employee.id !== editingId);
  if (duplicate) {
    alert("该员工姓名已存在，请换一个姓名或直接编辑原员工。");
    return;
  }

  if (editingId) {
    const previous = state.employees.find((employee) => employee.id === editingId);
    state.employees = state.employees.map((employee) =>
      employee.id === editingId
        ? {
            ...employee,
            name,
            team,
            monthlyTarget,
          }
        : employee,
    );

    if (previous && previous.name !== name) {
      for (const records of Object.values(state.attendance)) {
        if (records[previous.name]) {
          records[name] = records[previous.name];
          delete records[previous.name];
        }
      }

      for (const records of Object.values(state.performance)) {
        if (records[previous.name]) {
          records[name] = records[previous.name];
          delete records[previous.name];
        }
      }
    }
  } else {
    state.employees.push({
      id: createId(),
      name,
      team,
      dailyTarget: 0,
      monthlyTarget,
    });
  }

  resetEmployeeForm();
  syncAfterChange();
}

function startEditTeam(teamId) {
  const team = state.teams.find((item) => item.id === teamId);
  if (!team) {
    return;
  }

  els.teamEditingId.value = team.id;
  els.teamName.value = team.name;
  els.teamLeader.value = team.leader || "";
  els.teamDailyDeviceTarget.value = String(team.dailyDeviceTarget);
  els.teamMonthlyDeviceTarget.value = String(team.monthlyDeviceTarget);
  els.teamSubmitBtn.textContent = "保存团队";
}

function startEditEmployee(employeeId) {
  const employee = state.employees.find((item) => item.id === employeeId);
  if (!employee) {
    return;
  }

  els.employeeEditingId.value = employee.id;
  els.employeeName.value = employee.name;
  els.employeeTeam.value = employee.team;
  els.employeeMonthlyTarget.value = String(employee.monthlyTarget);
  els.employeeSubmitBtn.textContent = "保存员工";
}

function deleteTeam(teamId) {
  const team = state.teams.find((item) => item.id === teamId);
  if (!team) {
    return;
  }

  if (state.employees.some((employee) => employee.team === team.name)) {
    alert("该团队下还有员工，请先调整员工归属后再删除团队。");
    return;
  }

  state.teams = state.teams.filter((item) => item.id !== teamId);
  syncAfterChange();
}

function deleteEmployee(employeeId) {
  const employee = state.employees.find((item) => item.id === employeeId);
  if (!employee) {
    return;
  }

  state.employees = state.employees.filter((item) => item.id !== employeeId);

  for (const records of Object.values(state.attendance)) {
    delete records[employee.name];
  }

  for (const records of Object.values(state.performance)) {
    delete records[employee.name];
  }

  syncAfterChange();
}

function handleTableClick(event) {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }

  if (target.dataset.action === "delete-employee") {
    deleteEmployee(target.dataset.id);
  }

  if (target.dataset.action === "edit-employee") {
    startEditEmployee(target.dataset.id);
  }

  if (target.dataset.action === "delete-team") {
    deleteTeam(target.dataset.id);
  }

  if (target.dataset.action === "edit-team") {
    startEditTeam(target.dataset.id);
  }
}

function handleTableInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) {
    return;
  }

  if (target.dataset.action !== "performance-input") {
    return;
  }

  const dateKey = els.dateInput.value;
  const employeeName = target.dataset.name;
  const field = target.dataset.field;
  if (!employeeName || !field) {
    return;
  }

  ensureDailyPerformance(dateKey, employeeName)[field] = Math.max(0, Number(target.value || 0));
  scheduleSave({ render: false, immediate: false });
}

function handleTableChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) {
    return;
  }

  const dateKey = els.dateInput.value;
  const employeeName = target.dataset.name;
  if (!employeeName) {
    return;
  }

  if (target.dataset.action === "toggle-attendance") {
    ensureDailyAttendance(dateKey, employeeName)[target.dataset.slot] = target.checked;
    syncAfterChange();
  }

  if (target.dataset.action === "performance-input") {
    ensureDailyPerformance(dateKey, employeeName)[target.dataset.field] = Math.max(0, Number(target.value || 0));
    syncAfterChange();
  }
}

function buildExportRows() {
  const dateKey = els.dateInput.value;
  const employees = getFilteredEmployees();

  const employeeRows = employees.map((employee) => ({
    姓名: employee.name,
    团队: employee.team,
    团队长: getTeamProfile(employee.team).leader || "",
    月目标: employee.monthlyTarget,
  }));

  const attendanceRows = employees.map((employee, index) => {
    const record = ensureDailyAttendance(dateKey, employee.name);
    const row = {
      日期: dateKey,
      序号: index + 1,
      团队归属: employee.team,
      团队长: getTeamProfile(employee.team).leader || "",
      作业BD: employee.name,
    };

    for (const slot of TIME_SLOTS) {
      row[slot] = record[slot] ? "√" : "";
    }

    row.出勤次数 = calculateAttendanceCount(dateKey, employee.name);
    return row;
  });

  const performanceRows = employees.map((employee) => {
    const metrics = getEmployeeMetrics(dateKey, employee);

    return {
      日期: dateKey,
      姓名: employee.name,
      N5派单: Number(metrics.daily.assigned || 0),
      N5自拓: Number(metrics.daily.selfDeveloped || 0),
      今日量: metrics.todayAmount,
      当日日目标: metrics.dailyTarget,
      完成率: percent(metrics.dailyRatio),
      当月派单总量: metrics.monthly.assigned,
      当月自拓总量: metrics.monthly.selfDeveloped,
      当月达成: metrics.monthly.achieved,
      个人月目标: employee.monthlyTarget,
      月度完成率: percent(metrics.monthlyRatio),
    };
  });

  const summaryRows = employees.map((employee) => {
    const metrics = getEmployeeMetrics(dateKey, employee);
    const team = getTeamProfile(employee.team);

    return {
      日期: dateKey,
      团队: employee.team,
      团队长: team.leader || "",
      姓名: employee.name,
      出勤次数: metrics.attendanceCount,
      今日量: metrics.todayAmount,
      当日日目标: metrics.dailyTarget,
      今日完成率: percent(metrics.dailyRatio),
      月度达成: metrics.monthly.achieved,
      月度目标: employee.monthlyTarget,
      月度完成率: percent(metrics.monthlyRatio),
      团队日设备目标: team.dailyDeviceTarget || 0,
      团队月设备目标: team.monthlyDeviceTarget || 0,
      状态: metrics.status.text,
    };
  });

  return {
    employeeRows,
    attendanceRows,
    performanceRows,
    summaryRows,
  };
}

function applySheetColumnWidths(worksheet, keys) {
  worksheet["!cols"] = keys.map((key) => ({
    wch: Math.max(String(key).length * 2, 12),
  }));
}

function exportExcel() {
  if (!globalThis.XLSX) {
    alert("当前浏览器未成功加载 Excel 组件，请刷新页面后重试。");
    return;
  }

  const { employeeRows, attendanceRows, performanceRows, summaryRows } = buildExportRows();
  const workbook = XLSX.utils.book_new();
  const sheets = [
    { name: "员工档案", rows: employeeRows },
    { name: "出勤打卡", rows: attendanceRows },
    { name: "业绩统计", rows: performanceRows },
    { name: "整合看板", rows: summaryRows },
  ];

  for (const sheet of sheets) {
    const rows = sheet.rows.length ? sheet.rows : [{ 提示: "当前筛选条件下没有数据" }];
    const worksheet = XLSX.utils.json_to_sheet(rows);
    applySheetColumnWidths(worksheet, Object.keys(rows[0]));
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
  }

  XLSX.writeFile(workbook, `出勤业绩整合报表-${els.dateInput.value || getToday()}.xlsx`);
}

function exportData() {
  const blob = new Blob([JSON.stringify(normalizeState(state), null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `出勤业绩整合数据-${els.dateInput.value || getToday()}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
}

function importData(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      state = normalizeState(JSON.parse(String(reader.result)));
      syncAfterChange();
    } catch (error) {
      alert(`导入失败：${error.message}`);
    }
  };
  reader.readAsText(file, "utf-8");
}

function resetData() {
  state = getSeedData();
  resetTeamForm();
  resetEmployeeForm();
  els.dateInput.value = getToday();
  els.teamFilter.value = "";
  els.quickSearch.value = "";
  syncAfterChange();
}

function getSupabaseConfig() {
  const config = globalThis.SUPABASE_CONFIG ?? {};
  if (!config.url || !config.anonKey || !globalThis.supabase?.createClient) {
    return null;
  }

  return {
    url: config.url,
    anonKey: config.anonKey,
  };
}

async function clearCloudChannel() {
  if (!cloud.channel) {
    return;
  }

  try {
    if (cloud.client?.removeChannel) {
      await cloud.client.removeChannel(cloud.channel);
    } else if (cloud.channel.unsubscribe) {
      await cloud.channel.unsubscribe();
    }
  } catch (error) {
    console.warn("移除云端订阅失败", error);
  } finally {
    cloud.channel = null;
  }
}

async function hydrateRemoteState() {
  if (!cloud.client || !cloud.session) {
    return;
  }

  const { data, error } = await cloud.client
    .from(SUPABASE_TABLE)
    .select("payload, updated_at")
    .eq("user_id", cloud.session.user.id);

  if (error) {
    setAuthStatus(`读取云端数据失败：${error.message}`, "error");
    setCloudBadge("同步异常", "cloud-error");
    return;
  }

  if (Array.isArray(data) && data.length > 0 && data[0].payload) {
    state = normalizeState(data[0].payload);
    persistLocalState();
    cloud.signature = getStateSignature(state);
    cloud.hydrated = true;
    lastSavedAt = data[0].updated_at ? new Date(data[0].updated_at) : new Date();
    updateSaveStatus("saved");
    renderAll();
    setAuthStatus("已从云端恢复最新数据。", "success");
    return;
  }

  cloud.hydrated = true;
  await persistCloudState(true);
  updateSaveStatus("saved");
  setAuthStatus("当前账号还没有云端数据，已把本地数据初始化到云端。", "success");
}

async function subscribeToRemoteState() {
  if (!cloud.client || !cloud.session) {
    return;
  }

  await clearCloudChannel();
  cloud.channel = cloud.client
    .channel(`app-state-${cloud.session.user.id}`)
    .on(
      "postgres_changes",
      {
        event: "*",
        schema: "public",
        table: SUPABASE_TABLE,
        filter: `user_id=eq.${cloud.session.user.id}`,
      },
      (payload) => {
        const nextPayload = payload.new?.payload;
        if (!nextPayload) {
          return;
        }

        const normalized = normalizeState(nextPayload);
        const nextSignature = getStateSignature(normalized);
        if (nextSignature === getStateSignature(state)) {
          return;
        }

        state = normalized;
        cloud.signature = nextSignature;
        persistLocalState();
        lastSavedAt = payload.new?.updated_at ? new Date(payload.new.updated_at) : new Date();
        renderAll();
        updateSaveStatus("saved");
        setAuthStatus("检测到云端新数据，页面已自动同步。", "success");
      },
    )
    .subscribe();
}

async function applySession(session) {
  cloud.session = session ?? null;

  if (!cloud.session) {
    cloud.hydrated = false;
    cloud.signature = "";
    await clearCloudChannel();
    renderAuthPanel();
    updateSaveStatus("idle");
    return;
  }

  renderAuthPanel();
  setCloudBadge("正在同步", "cloud-warning");
  setAuthStatus("登录成功，正在读取云端数据...", "warning");
  await hydrateRemoteState();
  await subscribeToRemoteState();
  renderAuthPanel();
  setCloudBadge("云端在线", "cloud-online");
}

async function initCloud() {
  const config = getSupabaseConfig();
  cloud.configured = Boolean(config);
  renderAuthPanel();

  if (!config) {
    updateSaveStatus("idle");
    return;
  }

  cloud.client = globalThis.supabase.createClient(config.url, config.anonKey, {
    auth: {
      persistSession: true,
      autoRefreshToken: true,
      detectSessionInUrl: true,
    },
  });

  const { data, error } = await cloud.client.auth.getSession();
  if (error) {
    setAuthStatus(`读取登录状态失败：${error.message}`, "error");
  } else {
    await applySession(data.session);
  }

  const { data: authData } = cloud.client.auth.onAuthStateChange((event, session) => {
    window.setTimeout(() => {
      void applySession(session);
      if (event === "SIGNED_OUT") {
        setAuthStatus("已退出登录，页面切回本地保存模式。", "warning");
      }
    }, 0);
  });

  cloud.authSubscription = authData.subscription;
}

async function handleSignIn() {
  if (!cloud.client) {
    setAuthStatus("当前还没有配置 Supabase。", "error");
    return;
  }

  const email = els.authEmail.value.trim();
  const password = els.authPassword.value;
  if (!email || !password) {
    setAuthStatus("请先填写邮箱和密码。", "warning");
    return;
  }

  setAuthStatus("正在登录云端账号...", "warning");
  const { error } = await cloud.client.auth.signInWithPassword({
    email,
    password,
  });

  if (error) {
    setAuthStatus(`登录失败：${error.message}`, "error");
    return;
  }

  els.authPassword.value = "";
  setAuthStatus("登录成功，正在同步数据...", "success");
}

async function handleSignUp() {
  if (!cloud.client) {
    setAuthStatus("当前还没有配置 Supabase。", "error");
    return;
  }

  const email = els.authEmail.value.trim();
  const password = els.authPassword.value;
  if (!email || !password) {
    setAuthStatus("请先填写邮箱和密码。", "warning");
    return;
  }

  setAuthStatus("正在创建云端账号...", "warning");
  const { data, error } = await cloud.client.auth.signUp({
    email,
    password,
    options: {
      emailRedirectTo: window.location.href,
    },
  });

  if (error) {
    setAuthStatus(`注册失败：${error.message}`, "error");
    return;
  }

  els.authPassword.value = "";
  if (!data.session) {
    setAuthStatus("注册成功，请先去邮箱完成验证，再回来登录。", "success");
  } else {
    setAuthStatus("注册成功，账号已自动登录。", "success");
  }
}

async function handleSignOut() {
  if (!cloud.client) {
    return;
  }

  const { error } = await cloud.client.auth.signOut();
  if (error) {
    setAuthStatus(`退出失败：${error.message}`, "error");
    return;
  }

  setAuthStatus("已退出登录。", "warning");
}

function bindEvents() {
  els.teamForm.addEventListener("submit", addTeam);
  els.employeeForm.addEventListener("submit", addEmployee);
  els.signInBtn.addEventListener("click", () => {
    void handleSignIn();
  });
  els.signUpBtn.addEventListener("click", () => {
    void handleSignUp();
  });
  els.signOutBtn.addEventListener("click", () => {
    void handleSignOut();
  });
  document.body.addEventListener("click", handleTableClick);
  document.body.addEventListener("input", handleTableInput);
  document.body.addEventListener("change", handleTableChange);
  els.dateInput.addEventListener("change", renderAll);
  els.teamFilter.addEventListener("change", renderAll);
  els.quickSearch.addEventListener("input", renderAll);
  els.exportExcelBtn.addEventListener("click", exportExcel);
  els.exportBtn.addEventListener("click", exportData);
  els.importBtn.addEventListener("click", () => els.importInput.click());
  els.importInput.addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    if (file) {
      importData(file);
    }
    event.target.value = "";
  });
  els.resetBtn.addEventListener("click", resetData);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") {
      flushPendingSave();
    }
  });
  window.addEventListener("beforeunload", flushPendingSave);
}

async function init() {
  els.dateInput.value = getToday();
  bindEvents();
  renderAll();
  updateSaveStatus("idle");
  await initCloud();
}

void init();
