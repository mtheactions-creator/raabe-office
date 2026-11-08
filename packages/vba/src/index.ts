// @raabe/vba (D-079): macros for RaabePresent and RaabeCalc. The VBA project format of Office files (compound
// file, compression, modules, UserForms) read and written, and our interpreter of the language; the apps give it
// their object model. Macros reach outside the app only through the app and after asking the user (D-080).

export { readCfb, readCfbStorages, writeCfb } from "./cfb";
export { compress, decompress, readProject, visibleCode, writeProject, type VbaModule } from "./project";
export { NOT_A_METHOD, VbaError, VbaRuntime, parseModule, vbaToNumber, vbaToString, type VbaArray, type VbaHost, type VbaObject, type VbaValue } from "./interp";
export type { VbaBreak, VbaFileAccessKind, VbaFileSystem, VbaSettingsStore, VbaStackFrame, VbaStepMode, VbaVariable } from "./interp";
export { VbaCollection, VbaDictionary } from "./objects";
export { dateToSerial as vbaDateToSerial, serialToDate as vbaSerialToDate } from "./values";
export { readFormStorage, writeFormStorage } from "./oforms";
export { allControls, newControl, newForm, type VbaControlDesign, type VbaControlType, type VbaFont, type VbaForm, type VbaFormDesign, type VbaRawControl } from "./forms";
