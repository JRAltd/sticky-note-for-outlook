import React from "react";
import ReactDOM from "react-dom/client";
import { App } from "./App";
import "./styles/index.css";

function renderApp() {
  const rootEl = document.getElementById("root");
  if (!rootEl) return;
  ReactDOM.createRoot(rootEl).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
}

if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(() => {
    renderApp();
  });
} else {
  renderApp();
}
