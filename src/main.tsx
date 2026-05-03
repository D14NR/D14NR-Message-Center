import { StrictMode, useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { ThemeProvider, createTheme } from "@mui/material/styles";
import CssBaseline from "@mui/material/CssBaseline";
import "./index.css";
import App from "./App";

function Root() {
  const [mode, setMode] = useState<"light" | "dark">(() => (localStorage.getItem('app_theme') as "light" | "dark") || 'light');

  useEffect(() => {
    const onTheme = (e: Event) => {
      try {
        // CustomEvent carries detail
        // @ts-ignore
        const d = e.detail as "light" | "dark" | undefined;
        if (d) setMode(d);
      } catch (err) {}
    };
    window.addEventListener('theme-changed', onTheme as EventListener);
    return () => window.removeEventListener('theme-changed', onTheme as EventListener);
  }, []);

  const theme = createTheme({
    palette: {
      mode,
      primary: { 
        main: '#6366F1', // Indigo
        light: '#818CF8',
        dark: '#4F46E5',
        contrastText: '#fff',
      },
      secondary: { 
        main: '#8B5CF6', // Violet
        light: '#A78BFA',
        dark: '#7C3AED',
        contrastText: '#fff',
      },
      success: { main: '#10B981' }, // Emerald
      error: { main: '#EF4444' }, // Red
      warning: { main: '#F59E0B' }, // Amber
      info: { main: '#06B6D4' }, // Cyan
      background: {
        default: mode === 'light' ? '#F9FAFB' : '#0F172A',
        paper: mode === 'light' ? '#FFFFFF' : '#1E293B',
      },
      text: {
        primary: mode === 'light' ? '#1F2937' : '#F1F5F9',
        secondary: mode === 'light' ? '#6B7280' : '#CBD5E1',
      },
    },
    components: {
      MuiButton: { defaultProps: { disableElevation: true } },
      MuiPaper: {
        styleOverrides: {
          root: {
            backgroundImage: 'none',
          },
        },
      },
    },
    typography: {
      fontFamily: '"Inter", "Segoe UI", sans-serif',
      button: {
        textTransform: 'none',
        fontWeight: 600,
      },
    },
  });

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <App />
    </ThemeProvider>
  );
}

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <Root />
  </StrictMode>
);
