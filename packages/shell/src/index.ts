// The shared Office shell for RaabeWrite, RaabeCalc and RaabePresent.

import "./theme.css";

export { h, icon } from "./dom";
export { Commands, type Command } from "./commands";
export { Ribbon, type Tab, type Group, type Control, type Size } from "./ribbon";
export { TitleBar, type DocTab } from "./titlebar";
export { StatusBar, type StatusInfo } from "./statusbar";
export { Backstage, type BackstageEntry } from "./backstage";
export { Keytips } from "./keytips";
export { showDialog, messageBox, type DialogButton } from "./dialog";
export { showMenu, showPopup, closePopup, installTooltips, type MenuEntry } from "./popup";
export { showFontColors, showHighlights, highlightHex, HIGHLIGHTS } from "./palette";
export { baseName, dirName, type Platform, type RecoveryEntry, type FileInfo } from "./platform";
export { tauriPlatform } from "./platform-tauri";
export { browserPlatform } from "./platform-browser";

export function isTauri(): boolean {
  return "__TAURI_INTERNALS__" in window;
}
