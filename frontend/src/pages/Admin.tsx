/**
 * Admin page — HR management panel.
 * Tabs: 評分進度 | 系統設定 | 員工同步
 */

import { useState } from "react";
import { useNavigate } from "react-router-dom";
import { useApi } from "../hooks/useApi";
import { api } from "../services/api";
import type { Settings } from "../types";

type Tab = "progress" | "settings" | "employees";

export default function Admin() {
  const navigate = useNavigate();
  const [tab, setTab] = useState<Tab>("progress");

  return (
    <div className="page">
      <div className="header" style={{ cursor: "pointer" }} onClick={() => navigate("/")}>
        <div>
          <h1>📋 考核評分系統</h1>
          <div className="subtitle">← HR 管理後台</div>
        </div>
      </div>

      <div className="tab-bar">
        {(["progress", "settings", "employees"] as Tab[]).map((t) => (
          <button
            key={t}
            className={`tab-btn${tab === t ? " active" : ""}`}
            onClick={() => setTab(t)}
          >
            {tabLabel(t)}
          </button>
        ))}
      </div>

      {tab === "progress" && <ProgressTab />}
      {tab === "settings" && <SettingsTab />}
      {tab === "employees" && <EmployeesTab />}
    </div>
  );
}

// ── Progress tab ──────────────────────────────────────────────────────────

function ProgressTab() {
  const { data, loading, error } = useApi(
    () => api.get("/api/scoring/all-status").then((r) => r.data)
  );

  if (loading) return <div className="loading"><div className="spinner" />載入中...</div>;
  if (error) return <div className="error-page">{error}</div>;

  return (
    <div className="admin-section">
      <h3>各主管評分進度</h3>
      <div className="progress-list">
        {(data as any[]).map((m: any) => {
          const pct = m.total > 0 ? Math.round((m.scored / m.total) * 100) : 0;
          return (
            <div key={m.lineUid} className="progress-row">
              <div className="progress-name">{m.managerName}</div>
              <div className="progress-count">{m.scored}/{m.total}</div>
              <div className="progress-bar">
                <div className="progress-fill" style={{ width: `${pct}%` }} />
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── Settings tab ──────────────────────────────────────────────────────────

function SettingsTab() {
  const { data, loading, error, refetch } = useApi<Settings>(
    () => api.get("/api/admin/settings").then((r) => r.data)
  );
  const [edits, setEdits] = useState<Settings>({});
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState("");

  const EDITABLE_KEYS = [
    "當前季度",
    "評分期間描述",
    "評分開始日",
    "評分截止日",
    "試用期天數",
    "最低評分天數",
    "綁定驗證碼",
  ];

  function showToast(msg: string) {
    setToast(msg);
    setTimeout(() => setToast(""), 3000);
  }

  async function handleSave() {
    setSaving(true);
    try {
      await api.post("/api/admin/settings", edits);
      showToast("✅ 設定已更新");
      setEdits({});
      refetch();
    } catch (err: any) {
      showToast(`❌ ${err.message}`);
    } finally {
      setSaving(false);
    }
  }

  if (loading) return <div className="loading"><div className="spinner" />載入中...</div>;
  if (error) return <div className="error-page">{error}</div>;

  return (
    <div className="admin-section">
      <h3>系統設定</h3>
      {EDITABLE_KEYS.map((key) => (
        <div key={key} className="setting-row">
          <label>{key}</label>
          <input
            type="text"
            value={key in edits ? edits[key] : (data?.[key] ?? "")}
            onChange={(e) => setEdits((s) => ({ ...s, [key]: e.target.value }))}
          />
        </div>
      ))}
      <button className="btn-primary" onClick={handleSave} disabled={saving}>
        {saving ? "儲存中…" : "儲存設定"}
      </button>
      {toast && <div className="toast show">{toast}</div>}
    </div>
  );
}

// ── Employees tab ─────────────────────────────────────────────────────────

function EmployeesTab() {
  const { data, loading, error, refetch } = useApi(
    () => api.get("/api/admin/employees").then((r) => r.data)
  );
  const [syncing, setSyncing] = useState(false);
  const [toast, setToast] = useState("");

  function showToast(msg: string) {
    setToast(msg);
    setTimeout(() => setToast(""), 3000);
  }

  async function handleSync() {
    setSyncing(true);
    try {
      const { data: res } = await api.post("/api/admin/employees/sync");
      showToast(`✅ 同步完成，共 ${res.count} 位員工`);
      refetch();
    } catch (err: any) {
      showToast(`❌ ${err.message}`);
    } finally {
      setSyncing(false);
    }
  }

  if (loading) return <div className="loading"><div className="spinner" />載入中...</div>;
  if (error) return <div className="error-page">{error}</div>;

  return (
    <div className="admin-section">
      <div className="section-header">
        <h3>員工名單（{(data as any[]).length} 人）</h3>
        <button className="btn-secondary" onClick={handleSync} disabled={syncing}>
          {syncing ? "同步中…" : "🔄 從 HR 同步"}
        </button>
      </div>

      <table className="data-table">
        <thead>
          <tr>
            <th>員工編號</th>
            <th>姓名</th>
            <th>部門</th>
            <th>科別</th>
            <th>到職日</th>
          </tr>
        </thead>
        <tbody>
          {(data as any[]).map((emp: any) => (
            <tr key={emp.employeeId ?? emp.name}>
              <td>{emp.employeeId}</td>
              <td>{emp.name}</td>
              <td>{emp.dept}</td>
              <td>{emp.section}</td>
              <td>{emp.joinDate}</td>
            </tr>
          ))}
        </tbody>
      </table>

      {toast && <div className="toast show">{toast}</div>}
    </div>
  );
}

function tabLabel(t: Tab): string {
  return { progress: "評分進度", settings: "系統設定", employees: "員工名單" }[t];
}
