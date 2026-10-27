// RaabeCalc entry point.

import { browserPlatform, isTauri, tauriPlatform } from "@raabe/shell";
import "./calc.css";
import { CalcApp } from "./app";

const app = new CalcApp(isTauri() ? tauriPlatform() : browserPlatform());
const feel = (window as unknown as { __FEEL?: string }).__FEEL;
// The dev server and feel-lab runs expose the app for tests and the browser pane.
if (feel || import.meta.env.DEV) (window as unknown as { __calc: unknown }).__calc = app.testApi();
