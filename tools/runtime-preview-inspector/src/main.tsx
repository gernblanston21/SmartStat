import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import { adaptPreviewPayloadToViewModel } from "./adapters/previewPayloadToViewModel";
import "./styles.css";
import previewFixture from "../../../tests/_scratch/runtime-slice-02-readonly-plan-bridge/runs/pos_projection_intake_run1.json";

const viewModel = adaptPreviewPayloadToViewModel(previewFixture);

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App viewModel={viewModel} />
  </React.StrictMode>
);
