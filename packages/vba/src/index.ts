// @raabe/vba (D-079): macros for RaabePresent and RaabeCalc. The VBA project format of Office files (compound
// file, compression, modules) read and written, and our interpreter of the language; the apps give it their
// object model. Macros cannot reach outside the app (no Shell, files, CreateObject or Declare, D-066).

export { readCfb, writeCfb } from "./cfb";
export { compress, decompress, readProject, visibleCode, writeProject, type VbaModule } from "./project";
export { NOT_A_METHOD, VbaError, VbaRuntime, parseModule, vbaToNumber, vbaToString, type VbaArray, type VbaHost, type VbaObject, type VbaValue } from "./interp";
