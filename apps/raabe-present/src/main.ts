// RaabePresent entry point.

import { browserPlatform, isTauri, tauriPlatform } from "@raabe/shell";
import "./present.css";
import { PresentApp } from "./app";

const app = new PresentApp(isTauri() ? tauriPlatform() : browserPlatform());
const feel = (window as unknown as { __FEEL?: string }).__FEEL;
// The dev server and feel-lab runs expose the app for tests and the browser pane.
if (feel || import.meta.env.DEV) (window as unknown as { __present: unknown }).__present = app.testApi();
