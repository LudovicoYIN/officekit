export {
  getWordAdapterContract,
  getWordAdapterManifest,
  summarizeWordAdapterContract,
  summarizeWordAdapter,
  wordAdapterManifest
} from "./manifest.js";

export { getWordNode, queryWordNodes, getDocumentInfo, addWordNode, setWordNode, removeWordNode, moveWordNode, swapWordNodes, batchWordNodes, viewWordDocument, setWordStyle, setWordSection, setWordDocDefaults, rawWordDocument, rawSetWordDocument, setWordCompatibility, addWordPart, copyWordNode, ensureParaIds, setDocumentProperties, validateWordDocument, viewWordStatsJson, viewWordOutlineJson, viewWordTextJson, viewWordIssuesJson, getWordFormFields, setWordFormField, acceptAllTrackChanges, rejectAllTrackChanges, getDocumentProtection, setDocumentProtection, getWordSdts, setWordSdt } from "./adapter.js";
export { parsePath, buildPath, validatePath, isValidPath } from "./path.js";
export { parseSelector, buildSelector, validateSelector, isValidSelector } from "./selectors.js";
