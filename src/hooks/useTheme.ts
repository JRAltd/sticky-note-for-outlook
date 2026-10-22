import { useEffect, useState } from "react";
import { LOCAL_STORAGE_THEME_KEY } from "../config/constants";
import { ThemeMode } from "../types/note";

function detectSystemOrOutlookDark(): boolean {
  try {
    if (typeof Office !== "undefined" && Office.context?.officeTheme) {
      const bg = Office.context.officeTheme.bodyBackgroundColor;
      if (bg && (bg.startsWith("#0") || bg.startsWith("#1") || bg.startsWith("#2"))) {
        return true;
      }
    }
  } catch {
    // Fallback to prefers-color-scheme
  }

  return (
    typeof window !== "undefined" &&
    window.matchMedia &&
    window.matchMedia("(prefers-color-scheme: dark)").matches
  );
}

export function useTheme() {
  const [themeMode, setThemeMode] = useState<ThemeMode>(() => {
    const saved = localStorage.getItem(LOCAL_STORAGE_THEME_KEY) as ThemeMode | null;
    return saved === "light" || saved === "dark" || saved === "system"
      ? saved
      : "system";
  });

  const [isDark, setIsDark] = useState<boolean>(() => {
    const saved = localStorage.getItem(LOCAL_STORAGE_THEME_KEY);
    if (saved === "dark") return true;
    if (saved === "light") return false;
    return detectSystemOrOutlookDark();
  });

  useEffect(() => {
    localStorage.setItem(LOCAL_STORAGE_THEME_KEY, themeMode);
    if (themeMode === "dark") {
      setIsDark(true);
    } else if (themeMode === "light") {
      setIsDark(false);
    } else {
      setIsDark(detectSystemOrOutlookDark());
    }
  }, [themeMode]);

  useEffect(() => {
    document.documentElement.setAttribute(
      "data-theme",
      isDark ? "dark" : "light"
    );
  }, [isDark]);

  const cycleTheme = () => {
    setThemeMode((prev) => {
      if (prev === "system") return isDark ? "light" : "dark";
      if (prev === "light") return "dark";
      return "light";
    });
  };

  return {
    themeMode,
    setThemeMode,
    isDark,
    cycleTheme,
  };
}
