import { BrowserRouter, Navigate, Route, Routes } from "react-router-dom";
import { useLiff } from "./hooks/useLiff";
import Admin from "./pages/Admin";
import Bind from "./pages/Bind";
import Dashboard from "./pages/Dashboard";
import Score from "./pages/Score";
import SysAdmin from "./pages/SysAdmin";
import "./styles.css";

function AppRoutes() {
  const { ready, error } = useLiff();

  if (error) {
    return (
      <div className="page-center">
        <div className="card">
          <p className="error">⚠️ {error}</p>
          <button className="btn-primary" onClick={() => window.location.reload()}>
            重新整理
          </button>
        </div>
      </div>
    );
  }

  if (!ready) {
    return (
      <div className="page-center">
        <div className="spinner" />
        <p>初始化中...</p>
      </div>
    );
  }

  return (
    <Routes>
      <Route path="/" element={<Dashboard />} />
      <Route path="/bind" element={<Bind />} />
      <Route path="/score" element={<Score />} />
      <Route path="/admin" element={<Admin />} />
      <Route path="/sysadmin" element={<SysAdmin />} />
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}

export default function App() {
  return (
    <BrowserRouter>
      <AppRoutes />
    </BrowserRouter>
  );
}
