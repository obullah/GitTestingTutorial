import { useState } from "react";
import * as XLSX from "xlsx";
import {
  PieChart,
  Pie,
  Cell,
  Tooltip,
  Legend,
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
} from "recharts";
import "./App.css";

const PIE_COLORS = ["#22c55e", "#ef4444", "#eab308"];
const ROLE_COLORS = ["#6366f1", "#22c55e", "#f97316", "#ec4899", "#0ea5e9"];
const MEMBER_COLOR = "#38bdf8";

function App() {
  const [sprintName, setSprintName] = useState("");
  const [teamName, setTeamName] = useState("");

  // Sprint duration (dates)
  const [sprintStart, setSprintStart] = useState("");
  const [sprintEnd, setSprintEnd] = useState("");

  const [totalPerson, setTotalPerson] = useState("");
  const [productiveHrs, setProductiveHrs] = useState("");

  const [holidays, setHolidays] = useState("");
  const [regressionDays, setRegressionDays] = useState("");
  const [officialWorkingHrs, setOfficialWorkingHrs] = useState("8"); // default 8

  // Configurable roles
  const [roles, setRoles] = useState(["DEV", "QA"]);
  const [newRole, setNewRole] = useState("");

  const [members, setMembers] = useState([
    {
      id: 1,
      name: "",
      role: "DEV",
      leaves: "",
      hoursPerDay: "",
      workingFromClient: false,
      onshoreHolidays: "",
    },
  ]);

  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [showHelp, setShowHelp] = useState(false);

  // ---------- NEW: Jira Excel + API state ----------
  const [jiraIssues, setJiraIssues] = useState([]); // parsed issues from Excel
  const [jiraUploadError, setJiraUploadError] = useState("");
  const [jiraCreateLog, setJiraCreateLog] = useState(null);
  const [isCreatingJira, setIsCreatingJira] = useState(false);

  // Helper: calendar days between dates (inclusive)
  const getCalendarDays = () => {
    if (!sprintStart || !sprintEnd) return null;
    const start = new Date(sprintStart);
    const end = new Date(sprintEnd);
    if (isNaN(start.getTime()) || isNaN(end.getTime()) || end < start) {
      return null;
    }
    const diffMs = end.getTime() - start.getTime();
    const days = Math.floor(diffMs / (1000 * 60 * 60 * 24)) + 1; // inclusive
    return days;
  };

  // Helper: working days (Mon–Fri) between dates (inclusive)
  const getWorkingDaysExclWeekends = () => {
    if (!sprintStart || !sprintEnd) return null;
    const start = new Date(sprintStart);
    const end = new Date(sprintEnd);
    if (isNaN(start.getTime()) || isNaN(end.getTime()) || end < start) {
      return null;
    }

    let count = 0;
    const cur = new Date(start);
    while (cur <= end) {
      const day = cur.getDay(); // 0 = Sun, 6 = Sat
      if (day !== 0 && day !== 6) {
        count++;
      }
      cur.setDate(cur.getDate() + 1);
    }
    return count;
  };

  const calendarDays = getCalendarDays();
  const workingDays = getWorkingDaysExclWeekends();

  // -------- Roles handlers --------

  const handleAddRole = () => {
    const r = newRole.trim().toUpperCase();
    if (!r) return;
    if (roles.includes(r)) {
      setNewRole("");
      return;
    }
    setRoles((prev) => [...prev, r]);
    setNewRole("");
  };

  const handleRemoveRole = (role) => {
    // prevent removing roles currently in use
    const inUse = members.some((m) => m.role === role);
    if (inUse) {
      alert("Cannot remove a role that is currently assigned to a member.");
      return;
    }
    setRoles((prev) => prev.filter((r) => r !== role));
  };

  // -------- Member handlers --------

  const handleAddMember = () => {
    setMembers((prev) => [
      ...prev,
      {
        id: Date.now(),
        name: "",
        role: roles[0] || "DEV",
        leaves: "",
        hoursPerDay: "",
        workingFromClient: false,
        onshoreHolidays: "",
      },
    ]);
  };

  const handleMemberChange = (id, field, value) => {
    setMembers((prev) =>
      prev.map((m) => (m.id === id ? { ...m, [field]: value } : m))
    );
  };

  const handleMemberCheckboxChange = (id, field, checked) => {
    setMembers((prev) =>
      prev.map((m) => (m.id === id ? { ...m, [field]: checked } : m))
    );
  };

  const handleRemoveMember = (id) => {
    setMembers((prev) => prev.filter((m) => m.id !== id));
  };

  // -------- Excel helpers for members --------

  const parseBoolCell = (value) => {
    if (value === undefined || value === null) return false;
    const v = String(value).trim().toLowerCase();
    return v === "yes" || v === "true" || v === "y" || v === "1";
  };

  const handleDownloadMembersExcel = () => {
    const header = [
      [
        "Name",
        "Role",
        "HoursPerDay",
        "Leaves",
        "WorkingFromClient",
        "OnshoreHolidays",
      ],
    ];

    const rows =
      members.length > 0 && members.some((m) => m.name.trim() !== "")
        ? members.map((m) => [
            m.name,
            m.role,
            m.hoursPerDay,
            m.leaves,
            m.workingFromClient ? "Yes" : "No",
            m.onshoreHolidays,
          ])
        : [
            ["John Doe", "DEV", "6.5", "1", "No", "0"],
            ["Jane QA", "QA", "7", "0", "Yes", "2"],
          ];

    const ws = XLSX.utils.aoa_to_sheet([...header, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Members");
    XLSX.writeFile(wb, "team_members_template.xlsx");
  };

  const handleUploadMembersExcel = (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setError("");

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const ws = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

        const importedMembers = json
          .map((row, index) => {
            const name = row.Name || row.name || "";
            if (!name || String(name).trim() === "") return null;

            const roleValue = row.Role || row.role || roles[0] || "DEV";
            const normalizedRole = String(roleValue).toUpperCase();

            if (!roles.includes(normalizedRole)) {
              setRoles((prev) =>
                prev.includes(normalizedRole) ? prev : [...prev, normalizedRole]
              );
            }

            return {
              id: Date.now() + index,
              name: String(name),
              role: normalizedRole,
              hoursPerDay:
                row.HoursPerDay !== undefined && row.HoursPerDay !== null
                  ? String(row.HoursPerDay)
                  : "",
              leaves:
                row.Leaves !== undefined && row.Leaves !== null
                  ? String(row.Leaves)
                  : "",
              workingFromClient: parseBoolCell(
                row.WorkingFromClient ?? row.workingFromClient
              ),
              onshoreHolidays:
                row.OnshoreHolidays !== undefined &&
                row.OnshoreHolidays !== null
                  ? String(row.OnshoreHolidays)
                  : "",
            };
          })
          .filter(Boolean);

        if (importedMembers.length === 0) {
          setError(
            "No valid members found in the uploaded Excel. Please check the template."
          );
          return;
        }

        setMembers(importedMembers);
      } catch (err) {
        console.error(err);
        setError("Failed to parse Excel file. Please check the format.");
      }
    };
    reader.readAsArrayBuffer(file);

    // reset file input so same file can be re-selected if needed
    event.target.value = "";
  };

  // -------- NEW: Jira Excel helpers --------

  // normalize column name – trims and lowercases (no spaces)
  const normalizeHeader = (name) =>
    String(name || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "");

  const handleUploadJiraExcel = (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setJiraUploadError("");
    setJiraCreateLog(null);
    setJiraIssues([]);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const ws = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

        if (!json.length) {
          setJiraUploadError(
            "Excel is empty. Please provide at least one row of data."
          );
          return;
        }

        const parsed = json
          .map((row, index) => {
            // Try to read columns in a case-insensitive way
            const map = {};
            Object.keys(row).forEach((key) => {
              map[normalizeHeader(key)] = row[key];
            });

            const issueId = map["issueid"] || map["id"] || "";
            const assignee = map["assignee"] || "";
            const storyPoints =
              map["storypoints"] || map["storypoint"] || map["sp"] || "";
            const summary = map["summary"] || "";
            const description = map["description"] || "";

            // Skip rows that don't even have a summary
            if (!summary || String(summary).trim() === "") return null;

            return {
              excelRowIndex: index + 2, // +2 → include header row notation
              issueId: String(issueId || "").trim(),
              assignee: String(assignee || "").trim(),
              storyPoints: storyPoints !== "" ? Number(storyPoints) : null,
              summary: String(summary || "").trim(),
              description: String(description || "").trim(),
            };
          })
          .filter(Boolean);

        if (!parsed.length) {
          setJiraUploadError(
            "No valid rows found. Make sure at least 'Summary' column is filled."
          );
          return;
        }

        setJiraIssues(parsed);
      } catch (err) {
        console.error(err);
        setJiraUploadError(
          "Failed to parse Jira Excel file. Please check the format."
        );
      }
    };

    reader.readAsArrayBuffer(file);

    // allow re-select same file
    event.target.value = "";
  };

  const handleCreateJiraIssues = async () => {
    if (!jiraIssues.length) {
      setJiraUploadError("Please upload a Jira Excel file first.");
      return;
    }

    setIsCreatingJira(true);
    setJiraUploadError("");
    setJiraCreateLog(null);

    try {
      const response = await fetch("/api/jira/bulk-create", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ issues: jiraIssues }),
      });

      if (!response.ok) {
        const text = await response.text();
        throw new Error(
          `Jira API call failed (${response.status}). Details: ${text}`
        );
      }

      const data = await response.json();
      setJiraCreateLog(data);
    } catch (err) {
      console.error(err);
      setJiraUploadError(
        err.message || "Failed to create Jira issues. See console for details."
      );
    } finally {
      setIsCreatingJira(false);
    }
  };

  // -------- Calculate --------

  const handleCalculate = () => {
    setError("");
    setResult(null);

    const duration = getCalendarDays();
    const working = getWorkingDaysExclWeekends();
    if (!duration || !working) {
      setError("Please select a valid sprint start and end date.");
      return;
    }

    const p = Number(totalPerson);
    const hrs = Number(productiveHrs);
    const wd = working; // working days (Mon–Fri)
    const hd = Number(holidays) || 0;
    const reg = Number(regressionDays) || 0;
    const officialH = Number(officialWorkingHrs);

    // Basic validations
    if (!p || p <= 0) {
      setError("Please enter a valid Total Person (greater than 0).");
      return;
    }
    if (!hrs || hrs <= 0) {
      setError("Please enter valid Productive Hrs (greater than 0).");
      return;
    }
    if (hd < 0) {
      setError("Please enter valid Holidays within Sprint (0 or more).");
      return;
    }
    if (reg < 0) {
      setError("Please enter valid Days needed for Regression (0 or more).");
      return;
    }
    if (!officialH || officialH <= 0) {
      setError("Please enter a valid Official Working hrs (greater than 0).");
      return;
    }

    // Team-level effective days (uses global holidays & regression)
    const baseEffectiveDaysTeam = Math.max(0, wd - hd - reg);
    const overallEffectiveDays = baseEffectiveDaysTeam;

    // Overall capacity (using global hrs & total person)
    const overallTotalHours = p * hrs * overallEffectiveDays;
    const overallTotalCapacityDays = overallTotalHours / officialH;

    // Member-level calculations (ignore members with no name)
    const memberResults = members
      .filter((m) => m.name.trim() !== "")
      .map((m) => {
        const leaves = Number(m.leaves) || 0;
        const memberHrs = Number(m.hoursPerDay) || hrs; // override or fallback
        const onshore = Number(m.onshoreHolidays) || 0;

        // Start from working days minus regression and leaves
        const base = wd - reg - leaves;

        // If member works from client location, use on-shore holidays
        // and ignore the global "Holidays within Sprint"
        let effectiveDays;
        if (m.workingFromClient) {
          effectiveDays = Math.max(0, base - onshore);
        } else {
          effectiveDays = Math.max(0, base - hd);
        }

        const hours = effectiveDays * memberHrs;
        const days = hours / officialH;

        return {
          ...m,
          leaves,
          hoursPerDay: memberHrs,
          onshoreHolidays: onshore,
          effectiveDays,
          hours,
          days,
        };
      });

    // Role-level totals (for ALL roles)
    const roleTotals = {};
    memberResults.forEach((m) => {
      if (!roleTotals[m.role]) {
        roleTotals[m.role] = { hours: 0, days: 0 };
      }
      roleTotals[m.role].hours += m.hours;
      roleTotals[m.role].days += m.days;
    });

    // Final team capacity = sum of all role capacities
    let finalHours = 0;
    let finalDays = 0;
    Object.values(roleTotals).forEach((t) => {
      finalHours += t.hours;
      finalDays += t.days;
    });

    setResult({
      sprintName,
      teamName,
      sprintStart,
      sprintEnd,
      sprintDurationDays: duration, // calendar days
      sprintWorkingDays: wd, // working days Mon–Fri
      totalPerson: p,
      productiveHrs: hrs,
      holidays: hd,
      regressionDays: reg,
      officialWorkingHrs: officialH,
      overallEffectiveDays,
      overallTotalHours,
      overallTotalCapacityDays,
      members: memberResults,
      roleTotals,
      finalTeamTotals: { hours: finalHours, days: finalDays },
    });
  };

  const handleReset = () => {
    setSprintName("");
    setTeamName("");
    setSprintStart("");
    setSprintEnd("");
    setTotalPerson("");
    setProductiveHrs("");
    setHolidays("");
    setRegressionDays("");
    setOfficialWorkingHrs("8");
    setMembers([
      {
        id: 1,
        name: "",
        role: roles[0] || "DEV",
        leaves: "",
        hoursPerDay: "",
        workingFromClient: false,
        onshoreHolidays: "",
      },
    ]);
    setResult(null);
    setError("");
    setJiraIssues([]);
    setJiraUploadError("");
    setJiraCreateLog(null);
  };

  const handleExport = () => {
    if (!result) return;

    const {
      sprintName,
      teamName,
      sprintStart,
      sprintEnd,
      sprintDurationDays,
      sprintWorkingDays,
      totalPerson,
      productiveHrs,
      holidays,
      regressionDays,
      officialWorkingHrs,
      overallEffectiveDays,
      overallTotalHours,
      overallTotalCapacityDays,
      roleTotals,
      finalTeamTotals,
      members: memberResults,
    } = result;

    const data = [
      ["SPRINT CAPACITY REPORT"],
      ["Generated Date", new Date().toLocaleDateString()],
      [""],

      ["--- SPRINT DETAILS & INPUTS ---"],
      ["Sprint Name", sprintName || "-"],
      ["Team Name", teamName || "-"],
      ["Sprint Duration", `${sprintStart} to ${sprintEnd}`],
      ["Calendar Days", sprintDurationDays],
      ["Working Days (Mon-Fri)", sprintWorkingDays],
      [""],
      ["Total Persons", totalPerson],
      ["Productive Hrs/Day", productiveHrs],
      ["Global Holidays", holidays],
      ["Regression Days", regressionDays],
      ["Official Working Hrs", officialWorkingHrs],
      [""],

      ["--- OVERALL CAPACITY SUMMARY ---"],
      ["Team Effective Days", overallEffectiveDays],
      ["Total Capacity (Hours)", overallTotalHours],
      ["Total Capacity (Days/SP)", overallTotalCapacityDays.toFixed(2)],
      [""],

      ["--- CAPACITY BY ROLE ---"],
      ["Role", "Total Hours", "Total Days (SP)"],
    ];

    Object.entries(roleTotals).forEach(([role, t]) => {
      data.push([role, t.hours, t.days.toFixed(2)]);
    });

    data.push([
      "TOTAL TEAM",
      finalTeamTotals.hours,
      finalTeamTotals.days.toFixed(2),
    ]);
    data.push([""]);

    data.push(["--- DETAILED MEMBER BREAKDOWN ---"]);
    data.push([
      "Name",
      "Role",
      "Hrs/Day",
      "Leaves",
      "Client Loc",
      "Onshore Hols",
      "Eff. Days",
      "Capacity (Hrs)",
      "Capacity (Days)",
    ]);

    memberResults.forEach((m) => {
      data.push([
        m.name,
        m.role,
        m.hoursPerDay,
        m.leaves,
        m.workingFromClient ? "Yes" : "No",
        m.onshoreHolidays,
        m.effectiveDays,
        m.hours,
        m.days.toFixed(2),
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Capacity Report");

    const safeSprintName = (sprintName || "sprint")
      .replace(/\s+/g, "_")
      .toLowerCase();

    XLSX.writeFile(wb, `${safeSprintName}_capacity_report.xlsx`);
  };

  // ------- Chart data builders (for dashboard) -------

  const buildWorkingDaysPieData = () => {
    if (!result) return [];
    const effective = result.overallEffectiveDays || 0;
    const totalWorking = result.sprintWorkingDays || 0;
    const lost = Math.max(0, totalWorking - effective);
    if (totalWorking === 0) return [];
    return [
      { name: "Effective Days", value: effective },
      { name: "Lost (Holidays + Regression)", value: lost },
    ];
  };

  const buildRoleCapacityData = () => {
    if (!result) return [];
    return Object.entries(result.roleTotals || {}).map(([role, t]) => ({
      role,
      hours: t.hours,
      days: Number(t.days.toFixed(2)),
    }));
  };

  const buildMemberCapacityData = () => {
    if (!result) return [];
    return (result.members || []).map((m) => ({
      name: m.name,
      hours: m.hours,
      days: Number(m.days.toFixed(2)),
    }));
  };

  const workingPieData = buildWorkingDaysPieData();
  const roleCapacityData = buildRoleCapacityData();
  const memberCapacityData = buildMemberCapacityData();

  return (
    <div className="app">
      <div className="card">
        {/* HEADER BAR */}
        <div className="header-row">
          <div className="header-main">
            <div className="header-main-top">
              <h1 className="title">Capacity Calculation Tool</h1>
              <button
                type="button"
                className="help-button"
                onClick={() => setShowHelp(true)}
              >
                How to use
              </button>
            </div>
            <p className="subtitle">
              Calculate team sprint capacity{" "}
              <span className="highlight">
                (working days, regression, global & on-shore holidays)
              </span>
            </p>
          </div>

          <div className="header-meta">
            <div className="field header-field">
              <label>Sprint Name</label>
              <input
                type="text"
                value={sprintName}
                onChange={(e) => setSprintName(e.target.value)}
                placeholder="e.g. Sprint 25"
              />
            </div>
            <div className="field header-field">
              <label>Team Name</label>
              <input
                type="text"
                value={teamName}
                onChange={(e) => setTeamName(e.target.value)}
                placeholder="e.g. Phoenix"
              />
            </div>
          </div>
        </div>

        {/* TOP-LEVEL INPUTS */}
        <div className="form-row">
          <div className="field">
            <label>Total Person</label>
            <input
              type="number"
              value={totalPerson}
              onChange={(e) => setTotalPerson(e.target.value)}
              placeholder="e.g. 6"
              min="1"
            />
          </div>

          <div className="field">
            <label>Productive Hrs / day</label>
            <input
              type="number"
              value={productiveHrs}
              onChange={(e) => setProductiveHrs(e.target.value)}
              placeholder="e.g. 6.5"
              step="0.5"
              min="0"
            />
          </div>

          <div className="field field-sprint-duration">
            <label>Sprint Duration</label>
            <div className="date-range">
              <input
                type="date"
                value={sprintStart}
                onChange={(e) => setSprintStart(e.target.value)}
              />
              <span className="date-separator">to</span>
              <input
                type="date"
                value={sprintEnd}
                onChange={(e) => setSprintEnd(e.target.value)}
              />
            </div>
            {calendarDays && workingDays && (
              <div className="duration-hint">
                Duration: <strong>{calendarDays}</strong> days (
                <strong>{workingDays}</strong> working days, Mon–Fri)
              </div>
            )}
          </div>

          <div className="field">
            <label>Holidays within Sprint</label>
            <input
              type="number"
              value={holidays}
              onChange={(e) => setHolidays(e.target.value)}
              placeholder="e.g. 1"
              min="0"
            />
          </div>

          <div className="field">
            <label>Days needed for Regression</label>
            <input
              type="number"
              value={regressionDays}
              onChange={(e) => setRegressionDays(e.target.value)}
              placeholder="e.g. 2"
              min="0"
            />
          </div>

          <div className="field">
            <label>Official Working hrs</label>
            <input
              type="number"
              value={officialWorkingHrs}
              onChange={(e) => setOfficialWorkingHrs(e.target.value)}
              placeholder="e.g. 8"
              min="1"
              step="0.5"
            />
          </div>
        </div>

        {/* EXCEL IMPORT / EXPORT FOR MEMBERS */}
        <div className="import-export-section">
          <div className="import-export-header">
            <h2>Team Members via Excel</h2>
            <p className="import-hint">
              Download the template, edit in Excel, then upload it back to
              populate members.
            </p>
          </div>
          <div className="import-export-actions">
            <button type="button" onClick={handleDownloadMembersExcel}>
              Download Members Excel
            </button>

            <label className="upload-button">
              Upload Members Excel
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleUploadMembersExcel}
              />
            </label>
          </div>
          <p className="import-hint">
            Expected columns: <strong>Name</strong>, <strong>Role</strong>,{" "}
            <strong>HoursPerDay</strong>, <strong>Leaves</strong>,{" "}
            <strong>WorkingFromClient</strong>,{" "}
            <strong>OnshoreHolidays</strong>.
          </p>
        </div>

        {/* NEW: JIRA EXCEL IMPORT + CREATE SECTION */}
        <div className="jira-section">
          <div className="jira-header">
            <h2>Jira Stories via Excel</h2>
            <p className="jira-hint">
              Upload an Excel file with columns:{" "}
              <strong>Issue ID</strong>, <strong>Assignee</strong>,{" "}
              <strong>Story Points</strong>, <strong>Summary</strong>,{" "}
              <strong>Description</strong>. After upload, click &quot;Create
              Jira Issues&quot; to bulk-create stories.
            </p>
          </div>

          <div className="jira-actions">
            <label className="upload-button">
              Upload Jira Excel
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleUploadJiraExcel}
              />
            </label>

            <button
              type="button"
              onClick={handleCreateJiraIssues}
              disabled={!jiraIssues.length || isCreatingJira}
            >
              {isCreatingJira
                ? "Creating Jira issues..."
                : jiraIssues.length
                ? `Create ${jiraIssues.length} Jira Issues`
                : "Create Jira Issues"}
            </button>
          </div>

          {jiraIssues.length > 0 && (
            <div className="jira-preview">
              <div className="jira-preview-meta">
                <span>
                  <strong>{jiraIssues.length}</strong> issues parsed from Excel
                  and ready to create.
                </span>
                <span className="jira-preview-note">
                  Showing first 5 rows as preview.
                </span>
              </div>
              <div className="jira-preview-table-wrapper">
                <table className="jira-preview-table">
                  <thead>
                    <tr>
                      <th>#</th>
                      <th>Issue ID</th>
                      <th>Assignee</th>
                      <th>Story Points</th>
                      <th>Summary</th>
                      <th>Description</th>
                    </tr>
                  </thead>
                  <tbody>
                    {jiraIssues.slice(0, 5).map((issue, idx) => (
                      <tr key={idx}>
                        <td>{issue.excelRowIndex}</td>
                        <td>{issue.issueId}</td>
                        <td>{issue.assignee}</td>
                        <td>
                          {issue.storyPoints !== null
                            ? issue.storyPoints
                            : ""}
                        </td>
                        <td>{issue.summary}</td>
                        <td className="jira-description-cell">
                          {issue.description}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {jiraCreateLog && (
            <div className="jira-log">
              <h3>Jira Creation Result</h3>
              <p>
                Created:{" "}
                <strong>{jiraCreateLog.successCount ?? 0}</strong> | Failed:{" "}
                <strong>{jiraCreateLog.failureCount ?? 0}</strong>
              </p>
              {jiraCreateLog.results && jiraCreateLog.results.length > 0 && (
                <div className="jira-log-table-wrapper">
                  <table className="jira-log-table">
                    <thead>
                      <tr>
                        <th>Excel Row</th>
                        <th>Issue ID</th>
                        <th>Jira Key</th>
                        <th>Status</th>
                        <th>Error</th>
                      </tr>
                    </thead>
                    <tbody>
                      {jiraCreateLog.results.map((r, idx) => (
                        <tr key={idx}>
                          <td>{r.excelRowIndex ?? "-"}</td>
                          <td>{r.issueId || "-"}</td>
                          <td>{r.jiraKey || "-"}</td>
                          <td>{r.success ? "✅ Success" : "❌ Failed"}</td>
                          <td className="jira-log-error">
                            {r.errorMessage || "-"}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {jiraUploadError && (
            <div className="error jira-error">{jiraUploadError}</div>
          )}
        </div>

        {/* ROLES CONFIG SECTION */}
        <div className="roles-section">
          <div className="roles-header">
            <h2>Roles</h2>
            <div className="roles-add">
              <input
                type="text"
                value={newRole}
                onChange={(e) => setNewRole(e.target.value)}
                placeholder="e.g. PO, BA"
              />
              <button
                type="button"
                className="small-button"
                onClick={handleAddRole}
              >
                + Add Role
              </button>
            </div>
          </div>
          <div className="roles-chips">
            {roles.map((role) => (
              <span key={role} className="role-chip">
                {role}
                {roles.length > 1 && (
                  <button
                    type="button"
                    className="role-remove"
                    onClick={() => handleRemoveRole(role)}
                  >
                    ✕
                  </button>
                )}
              </span>
            ))}
          </div>
        </div>

        {/* TEAM MEMBERS SECTION */}
        <div className="members-section">
          <div className="members-header">
            <h2>Team Members</h2>
            <button
              type="button"
              className="small-button"
              onClick={handleAddMember}
            >
              + Add Member
            </button>
          </div>

          <div className="members-table">
            <div className="members-row members-row-header">
              <div>Member Name</div>
              <div>Role</div>
              <div>Hrs / day</div>
              <div>Leaves (days)</div>
              <div>Client Loc?</div>
              <div>On-shore Holidays</div>
              <div />
            </div>

            {members.map((member) => (
              <div key={member.id} className="members-row">
                <div>
                  <input
                    type="text"
                    value={member.name}
                    onChange={(e) =>
                      handleMemberChange(member.id, "name", e.target.value)
                    }
                    placeholder="e.g. John Doe"
                  />
                </div>
                <div>
                  <select
                    value={member.role}
                    onChange={(e) =>
                      handleMemberChange(member.id, "role", e.target.value)
                    }
                  >
                    {roles.map((role) => (
                      <option key={role} value={role}>
                        {role}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <input
                    type="number"
                    value={member.hoursPerDay}
                    onChange={(e) =>
                      handleMemberChange(
                        member.id,
                        "hoursPerDay",
                        e.target.value
                      )
                    }
                    placeholder="e.g. 6.5"
                    min="0"
                    step="0.5"
                  />
                </div>
                <div>
                  <input
                    type="number"
                    value={member.leaves}
                    onChange={(e) =>
                      handleMemberChange(member.id, "leaves", e.target.value)
                    }
                    placeholder="e.g. 1"
                    min="0"
                  />
                </div>
                <div style={{ textAlign: "center" }}>
                  <input
                    type="checkbox"
                    checked={member.workingFromClient || false}
                    onChange={(e) =>
                      handleMemberCheckboxChange(
                        member.id,
                        "workingFromClient",
                        e.target.checked
                      )
                    }
                    title="Working from client location"
                  />
                </div>
                <div>
                  <input
                    type="number"
                    value={member.onshoreHolidays}
                    onChange={(e) =>
                      handleMemberChange(
                        member.id,
                        "onshoreHolidays",
                        e.target.value
                      )
                    }
                    placeholder="e.g. 1"
                    min="0"
                  />
                </div>
                <div className="members-actions">
                  {members.length > 1 && (
                    <button
                      type="button"
                      className="icon-button"
                      onClick={() => handleRemoveMember(member.id)}
                    >
                      ✕
                    </button>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* BUTTONS */}
        <div className="buttons">
          <button onClick={handleCalculate}>Calculate</button>
          <button className="secondary" onClick={handleReset}>
            Reset
          </button>
          <button onClick={handleExport} disabled={!result}>
            Export Report
          </button>
        </div>

        {/* ERROR */}
        {error && <div className="error">{error}</div>}

        {/* RESULTS */}
        {result && (
          <>
            <div className="result">
              <div className="result-header">
                {result.sprintName && (
                  <span className="result-tag">
                    Sprint: <strong>{result.sprintName}</strong>
                  </span>
                )}
                {result.teamName && (
                  <span className="result-tag">
                    Team: <strong>{result.teamName}</strong>
                  </span>
                )}
                {result.sprintStart && result.sprintEnd && (
                  <span className="result-tag">
                    {result.sprintStart} → {result.sprintEnd}
                  </span>
                )}
              </div>

              <h2>Overall Capacity</h2>
              <div className="result-grid">
                <div className="result-card">
                  <span className="label">
                    Sprint Duration (calendar days)
                  </span>
                  <span className="value">{result.sprintDurationDays}</span>
                </div>
                <div className="result-card">
                  <span className="label">Working Days (Mon–Fri)</span>
                  <span className="value">{result.sprintWorkingDays}</span>
                </div>
                <div className="result-card">
                  <span className="label">Holidays (global)</span>
                  <span className="value">{result.holidays}</span>
                </div>
                <div className="result-card">
                  <span className="label">Regression Days</span>
                  <span className="value">{result.regressionDays}</span>
                </div>
                <div className="result-card">
                  <span className="label">Official Working hrs</span>
                  <span className="value">{result.officialWorkingHrs}</span>
                </div>
                <div className="result-card">
                  <span className="label">Effective Days (Overall)</span>
                  <span className="value">{result.overallEffectiveDays}</span>
                </div>
                <div className="result-card">
                  <span className="label">Total Capacity (hrs)</span>
                  <span className="value">{result.overallTotalHours}</span>
                </div>
                <div className="result-card">
                  <span className="label">Total Story Points</span>
                  <span className="value">
                    {result.overallTotalCapacityDays.toFixed(2)}
                  </span>
                </div>
              </div>

              <h2 className="subheading">Team Capacity by Role</h2>
              <div className="result-grid">
                {Object.entries(result.roleTotals).map(([role, t]) => (
                  <div className="result-card" key={role}>
                    <span className="label">{role} – Hours</span>
                    <span className="value">{t.hours}</span>
                    <span className="label" style={{ marginTop: "6px" }}>
                      {role} – Story Points
                    </span>
                    <span className="value">{t.days.toFixed(2)}</span>
                  </div>
                ))}
              </div>

              <h2 className="subheading">Final Team Capacity</h2>
              <div className="result-grid">
                <div className="result-card">
                  <span className="label">Total Hours (all roles)</span>
                  <span className="value">
                    {result.finalTeamTotals.hours}
                  </span>
                </div>
                <div className="result-card">
                  <span className="label">
                    Total Story Points (all roles)
                  </span>
                  <span className="value">
                    {result.finalTeamTotals.days.toFixed(2)}
                  </span>
                </div>
              </div>

              {result.members.length > 0 && (
                <>
                  <h2 className="subheading">Individual Capacity</h2>
                  <div className="members-result-table">
                    <div className="members-row members-row-header">
                      <div>Member</div>
                      <div>Role</div>
                      <div>Hrs / day</div>
                      <div>Leaves</div>
                      <div>Effective Days</div>
                      <div>Hours</div>
                      <div>Story Points</div>
                    </div>
                    {result.members.map((m) => (
                      <div key={m.id} className="members-row">
                        <div>{m.name}</div>
                        <div>{m.role}</div>
                        <div>{m.hoursPerDay}</div>
                        <div>{m.leaves}</div>
                        <div>{m.effectiveDays}</div>
                        <div>{m.hours}</div>
                        <div>{m.days.toFixed(2)}</div>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </div>

            {/* DASHBOARD SECTION WITH CHARTS */}
            <div className="dashboard-section">
              <h2 className="subheading">Visual Capacity Dashboard</h2>
              <div className="dashboard-grid">
                {/* Pie chart: Effective vs Lost days */}
                {workingPieData.length > 0 && (
                  <div className="chart-card">
                    <h3>Working Days Breakdown</h3>
                    <p className="chart-subtitle">
                      Effective vs lost (holidays & regression)
                    </p>
                    <div className="chart-inner">
                      <ResponsiveContainer width="100%" height={240}>
                        <PieChart>
                          <Pie
                            data={workingPieData}
                            dataKey="value"
                            nameKey="name"
                            innerRadius={50}
                            outerRadius={80}
                            paddingAngle={4}
                          >
                            {workingPieData.map((entry, index) => (
                              <Cell
                                key={`cell-${index}`}
                                fill={PIE_COLORS[index % PIE_COLORS.length]}
                              />
                            ))}
                          </Pie>
                          <Tooltip />
                          <Legend />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}

                {/* Bar chart: Role capacity */}
                {roleCapacityData.length > 0 && (
                  <div className="chart-card">
                    <h3>Capacity by Role</h3>
                    <p className="chart-subtitle">
                      Hours vs story-point days per role
                    </p>
                    <div className="chart-inner">
                      <ResponsiveContainer width="100%" height={260}>
                        <BarChart data={roleCapacityData}>
                          <CartesianGrid strokeDasharray="3 3" opacity={0.2} />
                          <XAxis dataKey="role" />
                          <YAxis />
                          <Tooltip />
                          <Legend />
                          <Bar dataKey="hours" name="Hours" radius={[6, 6, 0, 0]}>
                            {roleCapacityData.map((entry, index) => (
                              <Cell
                                key={`role-hours-${index}`}
                                fill={
                                  ROLE_COLORS[index % ROLE_COLORS.length]
                                }
                              />
                            ))}
                          </Bar>
                          <Bar
                            dataKey="days"
                            name="Days (SP)"
                            radius={[6, 6, 0, 0]}
                          />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}
              </div>

              {/* Member chart row */}
              {memberCapacityData.length > 0 && (
                <div className="dashboard-grid">
                  <div className="chart-card">
                    <h3>Individual Capacity (Hours)</h3>
                    <p className="chart-subtitle">
                      Each bar shows total capacity hours per member
                    </p>
                    <div className="chart-inner">
                      <ResponsiveContainer width="100%" height={280}>
                        <BarChart data={memberCapacityData}>
                          <CartesianGrid strokeDasharray="3 3" opacity={0.15} />
                          <XAxis
                            dataKey="name"
                            interval={0}
                            angle={-30}
                            textAnchor="end"
                            height={70}
                          />
                          <YAxis />
                          <Tooltip />
                          <Bar
                            dataKey="hours"
                            name="Hours"
                            radius={[6, 6, 0, 0]}
                            fill={MEMBER_COLOR}
                          />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </>
        )}
      </div>

      {/* HELP MODAL */}
      {showHelp && (
        <div className="help-overlay">
          <div className="help-modal">
            <div className="help-header">
              <h2>How to use the Capacity Calculation Tool</h2>
              <button
                type="button"
                className="help-close"
                onClick={() => setShowHelp(false)}
              >
                ✕
              </button>
            </div>
            <div className="help-content">
              {/* (Existing help content kept exactly as you had it) */}
              <p>
                This tool helps you calculate sprint capacity for your team in
                hours and “story point equivalent” days, considering: weekends,
                holidays, regression days, individual leaves and on-shore
                holidays.
              </p>

              <h3>1. Start with basic sprint information</h3>
              <ol>
                <li>
                  Enter <strong>Sprint Name</strong> and{" "}
                  <strong>Team Name</strong> for identification.
                </li>
                <li>
                  Set the <strong>Sprint Duration</strong> using{" "}
                  <strong>Start</strong> and <strong>End</strong> dates. The
                  tool automatically:
                  <ul>
                    <li>Counts total calendar days.</li>
                    <li>Counts only Monday–Friday as working days.</li>
                  </ul>
                </li>
                <li>
                  Fill in:
                  <ul>
                    <li>
                      <strong>Total Person</strong> – total number of people in
                      the team.
                    </li>
                    <li>
                      <strong>Productive Hrs / day</strong> – typical
                      productive hours per person (default for all members).
                    </li>
                    <li>
                      <strong>Holidays within Sprint</strong> – global
                      non-working days (for offshore team members).
                    </li>
                    <li>
                      <strong>Days needed for Regression</strong> – days
                      reserved for regression testing.
                    </li>
                    <li>
                      <strong>Official Working hrs</strong> – hours that
                      represent 1 working day (e.g. 8).
                    </li>
                  </ul>
                </li>
              </ol>

              <h3>2. Configure roles (DEV, QA, etc.)</h3>
              <ol>
                <li>
                  In the <strong>Roles</strong> section, keep default roles like
                  DEV / QA or add new ones (e.g. PO, BA).
                </li>
                <li>
                  You can remove a role only if no team member is currently
                  using it.
                </li>
              </ol>

              <h3>3. Add team members</h3>
              <p>You have two options:</p>
              <ul>
                <li>
                  <strong>Option A – Enter directly in UI</strong>
                  <ol>
                    <li>Click <strong>+ Add Member</strong>.</li>
                    <li>For each member, fill:</li>
                    <ul>
                      <li>
                        <strong>Member Name</strong>
                      </li>
                      <li>
                        <strong>Role</strong> – choose from the Roles list.
                      </li>
                      <li>
                        <strong>Hrs / day</strong> – if blank, default
                        Productive Hrs / day is used.
                      </li>
                      <li>
                        <strong>Leaves (days)</strong> – individual leave days.
                      </li>
                      <li>
                        <strong>Client Loc?</strong> – tick if the member works
                        from client location.
                      </li>
                      <li>
                        <strong>On-shore Holidays</strong> – holidays for this
                        member at client location (used only when Client Loc
                        is checked).
                      </li>
                    </ul>
                  </ol>
                </li>
                <li>
                  <strong>Option B – Use Excel (bulk upload)</strong>
                  <ol>
                    <li>
                      Click <strong>Download Members Excel</strong> to get a
                      template.
                    </li>
                    <li>
                      Open it in Excel and fill rows under these columns:
                      <ul>
                        <li>
                          <strong>Name</strong>
                        </li>
                        <li>
                          <strong>Role</strong> (e.g. DEV, QA)
                        </li>
                        <li>
                          <strong>HoursPerDay</strong>
                        </li>
                        <li>
                          <strong>Leaves</strong>
                        </li>
                        <li>
                          <strong>WorkingFromClient</strong> – Yes/No,
                          True/False or 1/0
                        </li>
                        <li>
                          <strong>OnshoreHolidays</strong>
                        </li>
                      </ul>
                    </li>
                    <li>
                      Save the file and click{" "}
                      <strong>Upload Members Excel</strong> to load members
                      into the UI.
                    </li>
                    <li>
                      After upload, you can still adjust values directly in the
                      UI if needed.
                    </li>
                  </ol>
                </li>
              </ul>

              <h3>4. How holidays are applied</h3>
              <ul>
                <li>
                  For <strong>offshore members</strong> (Client Loc not
                  checked), effective days use:
                  <br />
                  <em>
                    Working days – global Holidays within Sprint – Regression
                    days – individual Leaves
                  </em>
                </li>
                <li>
                  For <strong>client location members</strong> (Client Loc
                  checked), effective days use:
                  <br />
                  <em>
                    Working days – Regression days – individual Leaves – their
                    On-shore Holidays
                  </em>
                  <br />
                  Global holidays are ignored for them.
                </li>
              </ul>

              <h3>5. Calculate capacity</h3>
              <ol>
                <li>
                  Once sprint and member details are filled, click{" "}
                  <strong>Calculate</strong>.
                </li>
                <li>
                  If any required information is missing or invalid, an error
                  message will appear explaining what to fix.
                </li>
              </ol>

              <h3>6. Understand the results</h3>
              <ul>
                <li>
                  <strong>Overall Capacity</strong> shows:
                  <ul>
                    <li>Duration and working days.</li>
                    <li>Global holidays and regression days.</li>
                    <li>Effective days for the team as a whole.</li>
                    <li>Total capacity in hours and “story point” days.</li>
                  </ul>
                </li>
                <li>
                  <strong>Team Capacity by Role</strong> shows capacity for each
                  role (DEV, QA, etc.) in hours and days.
                </li>
                <li>
                  <strong>Final Team Capacity</strong> sums capacity of all
                  roles.
                </li>
                <li>
                  <strong>Individual Capacity</strong> shows each member’s:
                  <ul>
                    <li>Effective days</li>
                    <li>Capacity in hours</li>
                    <li>Capacity in story-point days</li>
                  </ul>
                </li>
              </ul>

              <h3>7. Export results</h3>
              <ol>
                <li>
                  Click <strong>Export Report</strong> to download an Excel file
                  with:
                  <ul>
                    <li>Sprint & team summary</li>
                    <li>Role-level totals</li>
                    <li>Final team capacity</li>
                    <li>Detailed member-level data</li>
                  </ul>
                </li>
                <li>
                  You can open this in Excel or any reporting tool for further
                  analysis.
                </li>
              </ol>

              <h3>8. Reset and start again</h3>
              <p>
                Click <strong>Reset</strong> at any time to clear values and
                start configuring a new sprint.
              </p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
