const STORAGE_KEY = "attendance-performance-system/v1";
const TIME_SLOTS = ["09:00", "10:00", "11:00", "12:00", "14:00", "15:00", "16:00", "17:00", "18:00"];

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
  employeeName: document.getElementById("employeeName"),
  employeeTeam: document.getElementById("employeeTeam"),
  employeeDailyTarget: document.getElementById("employeeDailyTarget"),
  employeeMonthlyTarget: document.getElementById("employeeMonthlyTarget"),
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
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportBtn: document.getElementById("exportBtn"),
  importBtn: document.getElementById("importBtn"),
  importInput: document.getElementById("importInput"),
  resetBtn: document.getElementById("resetBtn"),
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
    attendance: {
      [todayKey]: {
        丁宁: { "09:00": true, "10:00": false, "11:00": true, "12:00": true, "14:00": true, "15:00": false, "16:00": true, "17:00": true, "18:00": false },
        古雁和: { "09:00": true, "10:00": true, "11:00": false, "12:00": true, "14:00": true, "15:00": true, "16:00": false, "17:00": false, "18:00": false },
        江伟业: { "09:00": true, "10:00": true, "11:00": true, "12:00": true, "14:00": false, "15:00": true, "16:00": true, "17:00": false, "18:00": false },
        张超: { "09:00": false, "10:00": true, "11:00": true, "12:00": false, "14:00": true, "15:00": true, "16:00": true, "17:00": true, "18:00": false },
        邹春生: { "09:00": true, "10:00": true, "11:00": true, "12:00": false, "14:00": true, "15:00": false, "16:00": false, "17:00": true, "18:00": false },
      },
      [yesterdayKey]: {
        丁宁: { "09:00": true, "10:00": true, "11:00": true, "12:00": true, "14:00": true, "15:00": true, "16:00": false, "17:00": false, "18:00": false },
        古雁和: { "09:00": true, "10:00": true, "11:00": false, "12:00": true, "14:00": true, "15:00": false, "16:00": false, "17:00": false, "18:00": false },
      },
    },
    performance: {
      [todayKey]: {
        丁宁: { assigned: 1, selfDeveloped: 0 },
        古雁和: { assigned: 1, selfDeveloped: 0 },
        江伟业: { assigned: 2, selfDeveloped: 2 },
        张超: { assigned: 0, selfDeveloped: 2 },
        邹春生: { assigned: 1, selfDeveloped: 4 },
        郑文峰: { assigned: 0, selfDeveloped: 1 },
        谢烈雄: { assigned: 0, selfDeveloped: 0 },
        杨梓豪: { assigned: 0, selfDeveloped: 0 },
        曾浩: { assigned: 0, selfDeveloped: 1 },
      },
      [yesterdayKey]: {
        丁宁: { assigned: 1, selfDeveloped: 1 },
        古雁和: { assigned: 2, selfDeveloped: 0 },
        江伟业: { assigned: 1, selfDeveloped: 1 },
        张超: { assigned: 1, selfDeveloped: 0 },
        邹春生: { assigned: 0, selfDeveloped: 2 },
        郑文峰: { assigned: 0, selfDeveloped: 1 },
      },
    },
  };
}

function buildDerivedTeams(employees, existingTeams = []) {
  const totalsByTeam = new Map();

  for (const employee of employees) {
    if (!totalsByTeam.has(employee.team)) {
      totalsByTeam.set(employee.team, {
        id: createId(),
        name: employee.team,
        leader: "",
        dailyDeviceTarget: 0,
        monthlyDeviceTarget: 0,
      });
    }

    const team = totalsByTeam.get(employee.team);
    team.dailyDeviceTarget += Number(employee.dailyTarget || 0);
    team.monthlyDeviceTarget += Number(employee.monthlyTarget || 0);
  }

  for (const current of existingTeams) {
    const fallback = totalsByTeam.get(current.name) ?? {
      id: createId(),
      name: current.name,
      leader: "",
      dailyDeviceTarget: 0,
      monthlyDeviceTarget: 0,
    };

    totalsByTeam.set(current.name, {
      id: current.id || fallback.id,
      name: current.name,
      leader: current.leader ?? fallback.leader,
      dailyDeviceTarget: Number(current.dailyDeviceTarget ?? fallback.dailyDeviceTarget),
      monthlyDeviceTarget: Number(current.monthlyDeviceTarget ?? fallback.monthlyDeviceTarget),
    });
  }

  return [...totalsByTeam.values()].sort((a, b) => a.name.localeCompare(b.name, "zh-CN"));
}

function normalizeState(rawState) {
  const normalized = {
    teams: Array.isArray(rawState.teams) ? rawState.teams : [],
    employees: Array.isArray(rawState.employees) ? rawState.employees : [],
    attendance: rawState.attendance ?? {},
    performance: rawState.performance ?? {},
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
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed.employees)) {
      throw new Error("员工数据格式错误");
    }
    return normalizeState(parsed);
  } catch (error) {
    console.warn("读取本地数据失败，已恢复示例数据", error);
    return getSeedData();
  }
}

let state = loadState();

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
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
  state.performance[dateKey][employeeName] ??= { assigned: 0, selfDeveloped: 0 };
  state.performance[dateKey][employeeName].assigned ??= 0;
  state.performance[dateKey][employeeName].selfDeveloped ??= 0;
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
  const grouped = new Map();

  for (const employee of employees) {
    if (!grouped.has(employee.team)) {
      grouped.set(employee.team, []);
    }
    grouped.get(employee.team).push(employee);
  }

  return [...grouped.entries()]
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
  return `${(value * 100).toFixed(2)}%`;
}

function getEmployeeMetrics(dateKey, employee) {
  const attendanceCount = calculateAttendanceCount(dateKey, employee.name);
  const daily = ensureDailyPerformance(dateKey, employee.name);
  const todayAmount = Number(daily.assigned || 0) + Number(daily.selfDeveloped || 0);
  const dailyRatio = employee.dailyTarget ? todayAmount / employee.dailyTarget : 0;
  const monthly = calculateMonthlyPerformance(dateKey, employee.name);
  const monthlyRatio = employee.monthlyTarget ? monthly.achieved / employee.monthlyTarget : 0;
  const status = getStatusLabel(attendanceCount, dailyRatio, monthlyRatio);

  return {
    attendanceCount,
    daily,
    todayAmount,
    dailyRatio,
    monthly,
    monthlyRatio,
    status,
  };
}

function getTeamMetrics(dateKey, employees) {
  const totals = employees.reduce(
    (totals, employee) => {
      const { attendanceCount, todayAmount, monthly } = getEmployeeMetrics(dateKey, employee);
      totals.headcount += 1;
      totals.attendance += attendanceCount;
      totals.todayAmount += todayAmount;
      totals.monthlyAchieved += monthly.achieved;
      totals.monthlyTarget += employee.monthlyTarget;
      return totals;
    },
    { headcount: 0, attendance: 0, todayAmount: 0, monthlyAchieved: 0, monthlyTarget: 0 },
  );

  const teamProfile = employees[0] ? getTeamProfile(employees[0].team) : null;
  const dailyDeviceTarget = Number(teamProfile?.dailyDeviceTarget || 0);
  const monthlyDeviceTarget = Number(teamProfile?.monthlyDeviceTarget || 0);

  return {
    ...totals,
    teamProfile,
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
  const current = els.employeeTeam.value;
  const teams = uniqueTeams();
  els.employeeTeam.innerHTML = ['<option value="">请选择团队</option>', ...teams.map((team) => `<option value="${team}">${team}</option>`)].join("");
  els.employeeTeam.value = teams.includes(current) ? current : "";
}

function renderLegend() {
  els.slotLegend.innerHTML = TIME_SLOTS.map((slot) => `<span class="chip">${slot}</span>`).join("");
}

function renderTeams() {
  if (!state.teams.length) {
    els.teamTableBody.innerHTML = '<tr class="empty-row"><td colspan="5">当前还没有团队档案。</td></tr>';
    return;
  }

  els.teamTableBody.innerHTML = state.teams
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
    els.employeeTableBody.innerHTML = '<tr class="empty-row"><td colspan="5">当前筛选条件下没有员工。</td></tr>';
    return;
  }

  els.employeeTableBody.innerHTML = teamGroups
    .flatMap(({ team, members }) => [
      `<tr class="group-row"><td colspan="5">${team} · 共 ${members.length} 人 · 团队长 ${getTeamProfile(team).leader || "-"}</td></tr>`,
      ...members.map(
        (employee) => `
          <tr>
            <td>${employee.name}</td>
            <td>${employee.team}</td>
            <td>${employee.dailyTarget}</td>
            <td>${employee.monthlyTarget}</td>
            <td><button class="tiny-button danger" data-action="delete-employee" data-id="${employee.id}">删除</button></td>
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
    els.attendanceBody.innerHTML = '<tr class="empty-row"><td colspan="13">没有可显示的出勤数据。</td></tr>';
    return;
  }

  let index = 0;
  els.attendanceBody.innerHTML = teamGroups
    .flatMap(({ team, members }) => {
      const teamTotals = getTeamMetrics(dateKey, members);
      const subtotalRow = `<tr class="group-row"><td colspan="${TIME_SLOTS.length + 4}">${team} · 团队长 ${teamTotals.teamProfile?.leader || "-"} · ${members.length}人 · 团队出勤合计 ${teamTotals.attendance}</td></tr>`;
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
        const { daily, todayAmount, dailyRatio, monthly, monthlyRatio } = getEmployeeMetrics(dateKey, employee);

        return `
          <tr>
            <td>${employee.name}</td>
            <td>
              <input
                type="number"
                min="0"
                step="1"
                data-action="performance-input"
                data-field="assigned"
                data-name="${employee.name}"
                value="${daily.assigned}"
              />
            </td>
            <td>
              <input
                type="number"
                min="0"
                step="1"
                data-action="performance-input"
                data-field="selfDeveloped"
                data-name="${employee.name}"
                value="${daily.selfDeveloped}"
              />
            </td>
            <td>${todayAmount}</td>
            <td>${employee.dailyTarget}</td>
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

    const teamDailyRatio = percent(teamTotals.dailyDeviceCompletion);
    const teamMonthlyRatio = percent(teamTotals.monthlyDeviceCompletion);
    const subtotalRow = `
      <tr class="group-row">
        <td>${dateKey}</td>
        <td>${team}</td>
        <td>团队小计 / ${teamTotals.teamProfile?.leader || "-"}</td>
        <td>${teamTotals.attendance}</td>
        <td>${teamTotals.todayAmount}</td>
        <td>${teamTotals.dailyDeviceTarget}</td>
        <td>${teamDailyRatio}</td>
        <td>${teamTotals.monthlyAchieved}</td>
        <td>${teamTotals.monthlyDeviceTarget}</td>
        <td>${teamMonthlyRatio}</td>
        <td>团队汇总</td>
      </tr>
    `;

    const memberRows = members.map((employee) => {
      const { attendanceCount, todayAmount, dailyRatio, monthly, monthlyRatio, status } = getEmployeeMetrics(dateKey, employee);

      return `
        <tr>
          <td>${dateKey}</td>
          <td>${employee.team}</td>
          <td>${employee.name}</td>
          <td>${attendanceCount}</td>
          <td>${todayAmount}</td>
          <td>${employee.dailyTarget}</td>
          <td class="${getCompletionClass(dailyRatio)}">${percent(dailyRatio)}</td>
          <td>${monthly.achieved}</td>
          <td>${employee.monthlyTarget}</td>
          <td class="${getCompletionClass(monthlyRatio)}">${percent(monthlyRatio)}</td>
          <td><span class="status-pill ${status.className}">${status.text}</span></td>
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
    els.teamBoard.innerHTML = '<div class="team-card"><p>当前筛选条件下没有团队数据。</p></div>';
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
            <div class="team-metric">
              <span>团队人数</span>
              <strong>${totals.headcount}</strong>
            </div>
            <div class="team-metric">
              <span>当日出勤</span>
              <strong>${totals.attendance}</strong>
            </div>
            <div class="team-metric">
              <span>今日完成量</span>
              <strong>${totals.todayAmount}</strong>
            </div>
            <div class="team-metric">
              <span>日设备目标</span>
              <strong>${totals.dailyDeviceTarget}</strong>
            </div>
            <div class="team-metric">
              <span>日完成率</span>
              <strong>${percent(totals.dailyDeviceCompletion)}</strong>
            </div>
            <div class="team-metric">
              <span>月设备目标</span>
              <strong>${totals.monthlyDeviceTarget}</strong>
            </div>
            <div class="team-metric">
              <span>月完成率</span>
              <strong>${percent(totals.monthlyDeviceCompletion)}</strong>
            </div>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderAll() {
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
  saveState();
  renderAll();
}

function resetTeamForm() {
  els.teamForm.reset();
  els.teamEditingId.value = "";
  els.teamSubmitBtn.textContent = "新增团队";
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
  const name = els.employeeName.value.trim();
  const team = els.employeeTeam.value.trim();
  const dailyTarget = Number(els.employeeDailyTarget.value);
  const monthlyTarget = Number(els.employeeMonthlyTarget.value);

  if (!name || !team) {
    return;
  }

  const exists = state.employees.some((employee) => employee.name === name);
  if (exists) {
    alert("该员工已存在，请直接修改本地数据或先删除后重建。");
    return;
  }

  state.employees.push({
    id: createId(),
    name,
    team,
    dailyTarget,
    monthlyTarget,
  });

  els.employeeForm.reset();
  syncAfterChange();
}

function deleteTeam(teamId) {
  const team = state.teams.find((item) => item.id === teamId);
  if (!team) {
    return;
  }

  const linkedEmployees = state.employees.filter((employee) => employee.team === team.name);
  if (linkedEmployees.length) {
    alert("该团队下还有员工，先调整员工归属后再删除团队。");
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

  if (target.dataset.action === "delete-team") {
    deleteTeam(target.dataset.id);
  }

  if (target.dataset.action === "edit-team") {
    startEditTeam(target.dataset.id);
  }
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
    const slot = target.dataset.slot;
    ensureDailyAttendance(dateKey, employeeName)[slot] = target.checked;
    syncAfterChange();
  }

  if (target.dataset.action === "performance-input") {
    const field = target.dataset.field;
    const value = Math.max(0, Number(target.value || 0));
    ensureDailyPerformance(dateKey, employeeName)[field] = value;
    syncAfterChange();
  }
}

function exportData() {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `出勤业绩整合数据-${els.dateInput.value || getToday()}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
}

function buildExportRows() {
  const dateKey = els.dateInput.value;
  const employees = getFilteredEmployees();

  const employeeRows = employees.map((employee) => ({
    姓名: employee.name,
    团队: employee.team,
    团队长: getTeamProfile(employee.team).leader || "",
    日目标: employee.dailyTarget,
    月目标: employee.monthlyTarget,
  }));

  const attendanceRows = employees.map((employee, index) => {
    const record = ensureDailyAttendance(dateKey, employee.name);
    const row = {
      日期: dateKey,
      序号: index + 1,
      团队归属: employee.team,
      作业BD: employee.name,
    };

    for (const slot of TIME_SLOTS) {
      row[slot] = record[slot] ? "√" : "";
    }

    row.出勤次数 = calculateAttendanceCount(dateKey, employee.name);
    return row;
  });

  const performanceRows = employees.map((employee) => {
    const { daily, todayAmount, dailyRatio, monthly, monthlyRatio } = getEmployeeMetrics(dateKey, employee);

    return {
      日期: dateKey,
      姓名: employee.name,
      N5派单: Number(daily.assigned || 0),
      N5自拓: Number(daily.selfDeveloped || 0),
      今日量: todayAmount,
      个人日目标: employee.dailyTarget,
      完成率: percent(dailyRatio),
      当月派单总量: monthly.assigned,
      当月自拓总量: monthly.selfDeveloped,
      当月达成: monthly.achieved,
      月度目标: employee.monthlyTarget,
      月度完成率: percent(monthlyRatio),
    };
  });

  const summaryRows = employees.map((employee) => {
    const { attendanceCount, todayAmount, dailyRatio, monthly, monthlyRatio, status } = getEmployeeMetrics(dateKey, employee);

    return {
      日期: dateKey,
      团队: employee.team,
      团队长: getTeamProfile(employee.team).leader || "",
      姓名: employee.name,
      出勤次数: attendanceCount,
      今日量: todayAmount,
      日目标: employee.dailyTarget,
      今日完成率: percent(dailyRatio),
      月度达成: monthly.achieved,
      月度目标: employee.monthlyTarget,
      月度完成率: percent(monthlyRatio),
      团队日设备目标: getTeamProfile(employee.team).dailyDeviceTarget || 0,
      团队月设备目标: getTeamProfile(employee.team).monthlyDeviceTarget || 0,
      状态: status.text,
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
    alert("当前浏览器未成功加载 Excel 导出组件，请刷新页面后重试。");
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
    const worksheet = XLSX.utils.json_to_sheet(sheet.rows);
    const firstRow = sheet.rows[0] ?? {};
    applySheetColumnWidths(worksheet, Object.keys(firstRow));
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
  }

  XLSX.writeFile(workbook, `出勤业绩整合报表-${els.dateInput.value || getToday()}.xlsx`);
}

function importData(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result));
      if (!Array.isArray(parsed.employees) || typeof parsed.attendance !== "object" || typeof parsed.performance !== "object") {
        throw new Error("文件格式不正确");
      }
      state = normalizeState(parsed);
      syncAfterChange();
    } catch (error) {
      alert(`导入失败：${error.message}`);
    }
  };
  reader.readAsText(file, "utf-8");
}

function resetData() {
  state = getSeedData();
  els.dateInput.value = getToday();
  els.teamFilter.value = "";
  els.quickSearch.value = "";
  syncAfterChange();
}

function bindEvents() {
  els.teamForm.addEventListener("submit", addTeam);
  els.employeeForm.addEventListener("submit", addEmployee);
  document.body.addEventListener("click", handleTableClick);
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
}

function init() {
  els.dateInput.value = getToday();
  bindEvents();
  renderAll();
}

init();
