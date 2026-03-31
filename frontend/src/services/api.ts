/**
 * Axios client pre-configured for the Flask backend.
 * Automatically injects the session JWT from localStorage.
 */

import axios from "axios";

const BASE_URL = import.meta.env.VITE_API_URL ?? "http://localhost:8080";

export const api = axios.create({
  baseURL: BASE_URL,
  timeout: 30_000,
});

// Inject JWT on every request
api.interceptors.request.use((req) => {
  const token = localStorage.getItem("session_token");
  if (token) {
    req.headers.Authorization = `Bearer ${token}`;
  }
  return req;
});

// Normalise errors: always throw { error: string }
api.interceptors.response.use(
  (res) => res,
  (err) => {
    const serverMsg = err.response?.data?.error;
    throw new Error(serverMsg ?? err.message ?? "網路錯誤");
  }
);
