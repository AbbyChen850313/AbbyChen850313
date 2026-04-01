/**
 * LIFF initialisation hook.
 *
 * On mount:
 *   1. Check for an existing valid session JWT
 *   2. If none, delegate LIFF init/login to liffAdapter
 *   3. Exchange the LIFF access token for a session JWT
 *   4. Store JWT in localStorage and expose session info
 */

import { useCallback, useEffect, useState } from "react";
import { liffAdapter } from "../adapters/liff";
import { api } from "../services/api";

const IS_TEST = import.meta.env.VITE_IS_TEST === "true";
const SESSION_KEY = IS_TEST ? "session_token_test" : "session_token";

export interface LiffState {
  ready: boolean;
  needBind: boolean;
  error: string | null;
  lineUid: string | null;
  name: string | null;
  role: string | null;
}

export function useLiff(): LiffState {
  const [state, setState] = useState<LiffState>({
    ready: false,
    needBind: false,
    error: null,
    lineUid: null,
    name: null,
    role: null,
  });

  const initialise = useCallback(async () => {
    try {
      // 1. Check for existing valid session
      const existing = localStorage.getItem(SESSION_KEY);
      if (existing) {
        const { data } = await api.get("/api/auth/check", {
          headers: { Authorization: `Bearer ${existing}` },
        });
        setState({ ready: true, needBind: false, error: null, lineUid: null, name: data.name, role: data.role });
        return;
      }

      // 2. Fresh LIFF init via adapter (load SDK + liff.init + timeout)
      await liffAdapter.init();

      if (!liffAdapter.isLoggedIn()) {
        liffAdapter.login();
        return;
      }

      const accessToken = liffAdapter.getAccessToken();

      // 3. Exchange access token for session JWT
      const { data } = await api.post("/api/auth/session", {
        accessToken,
        isTest: IS_TEST,
      });

      localStorage.setItem(SESSION_KEY, data.token);

      setState({
        ready: true,
        needBind: false,
        error: null,
        lineUid: null,
        name: data.name,
        role: data.role,
      });
    } catch (err: any) {
      const msg: string = err?.message ?? "初始化失敗";

      // Account not bound → signal to render <Navigate to="/bind">
      if (err?.response?.data?.needBind || err?.message === "帳號未綁定") {
        setState((prev) => ({ ...prev, needBind: true }));
        return;
      }

      setState((prev) => ({ ...prev, ready: false, error: msg }));
    }
  }, []);

  useEffect(() => {
    initialise();
  }, [initialise]);

  return state;
}

/** Retrieve the stored session token (used by API service). */
export function getSessionToken(): string | null {
  return localStorage.getItem(SESSION_KEY);
}

/** Clear session and reload for logout. */
export function logout(): void {
  localStorage.removeItem(SESSION_KEY);
  liffAdapter.logout();
  window.location.reload();
}
