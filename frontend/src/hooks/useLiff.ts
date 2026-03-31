/**
 * LIFF initialisation hook.
 *
 * On mount:
 *   1. Load the LIFF SDK (injected as a script tag so it can be lazy-loaded)
 *   2. Call liff.init()
 *   3. If not logged in → liff.login()
 *   4. Verify the access token with our backend → receive session JWT
 *   5. Store JWT in localStorage and expose session info
 */

import { useCallback, useEffect, useState } from "react";
import { api } from "../services/api";

// Determine environment from LIFF ID env var
const LIFF_ID = import.meta.env.VITE_LIFF_ID as string;
const IS_TEST = import.meta.env.VITE_IS_TEST === "true";
const SESSION_KEY = IS_TEST ? "session_token_test" : "session_token";

declare global {
  interface Window {
    liff: any;
  }
}

export interface LiffState {
  ready: boolean;
  error: string | null;
  lineUid: string | null;
  name: string | null;
  role: string | null;
}

function loadLiffSdk(): Promise<void> {
  return new Promise((resolve, reject) => {
    if (window.liff) {
      resolve();
      return;
    }
    const script = document.createElement("script");
    script.src = "https://static.line-scdn.net/liff/edge/2/sdk.js";
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("LINE SDK 載入失敗"));
    document.head.appendChild(script);
  });
}

export function useLiff(): LiffState {
  const [state, setState] = useState<LiffState>({
    ready: false,
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
        setState({ ready: true, error: null, lineUid: null, name: data.name, role: data.role });
        return;
      }

      // 2. Fresh LIFF init
      await loadLiffSdk();
      await window.liff.init({ liffId: LIFF_ID });
      if (!window.liff.isLoggedIn()) {
        window.liff.login();
        return;
      }

      const accessToken: string = window.liff.getAccessToken();

      // 3. Exchange access token for session JWT
      const { data } = await api.post("/api/auth/session", {
        accessToken,
        isTest: IS_TEST,
      });

      localStorage.setItem(SESSION_KEY, data.token);

      setState({
        ready: true,
        error: null,
        lineUid: null,
        name: data.name,
        role: data.role,
      });
    } catch (err: any) {
      const msg: string = err?.message ?? "初始化失敗";

      // Account not bound → redirect to bind page
      if (err?.response?.data?.needBind) {
        window.location.href = "/bind";
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
  if (window.liff?.isLoggedIn?.()) {
    window.liff.logout();
  }
  window.location.reload();
}
