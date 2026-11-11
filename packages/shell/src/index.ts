// The shared Office shell for RaabeWrite, RaabeCalc and RaabePresent.

import "./theme.css";

export { h, icon } from "./dom";
export { Commands, type Command, type HoverTarget } from "./commands";
export { Ribbon, type Tab, type Group, type Control, type Size } from "./ribbon";
export { TitleBar, type DocTab } from "./titlebar";
export { StatusBar, type StatusInfo } from "./statusbar";
export { Backstage, type BackstageEntry } from "./backstage";
export { Keytips } from "./keytips";
export { installLostMouseUp } from "./drag";
export { showDialog, messageBox, type DialogButton } from "./dialog";
export { colorDialog, installDialogDrag } from "./color-dialog";
export { showMenu, showPopup, closePopup, hoverPreview, installTooltips, pressOutside, swallowPress, FRAME_SELECTOR, type MenuEntry } from "./popup";
export { TaskPane, type PaneSpec, type PaneSet, type PaneTab, type PaneSection, type PaneControl } from "./taskpane";
export { entryKind, pictureDropdown, pictureMenu, weightPicture, type PictureEntry, type PictureDropSpec, type PictureLive } from "./picturedrop";
export * as pictures from "./pictures";
export { ARROW_STYLES, arrowPairPicture, arrowStyleMenu, dashMenu, weightMenu } from "./linemenus";
export { colorGrid, showFontColors, showHighlights, highlightHex, HIGHLIGHTS, type PaletteHover, type PaletteTheme } from "./palette";
export { baseName, dirName, type Platform, type RecoveryEntry, type FileInfo, type GrammarIssue, type ThesaurusResult } from "./platform";
export { tauriPlatform } from "./platform-tauri";
export { browserPlatform } from "./platform-browser";

export function isTauri(): boolean {
  return "__TAURI_INTERNALS__" in window;
}
export { OFFICE_DESIGNS, OFFICE_COLOR_SETS, OFFICE_FONT_SETS } from "./office-designs.gen";
