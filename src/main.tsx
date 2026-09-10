import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import SlitterApp from "./slitter-app";
import "./globals.css";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <SlitterApp />
  </StrictMode>,
);
