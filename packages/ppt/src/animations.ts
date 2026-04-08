/**
 * Animation operations for @officekit/ppt.
 *
 * Provides functions to manage animations on slide elements:
 * - getAnimations: Get all animations on a slide
 * - addAnimation: Add an animation to a shape
 * - removeAnimation: Remove an animation from a shape
 *
 * Animation types: entrance, exit, emphasis, motionPath
 * Animation triggers: onClick, afterPrev, withPrev, onLoad
 *
 * Reference: OfficeCLI PowerPointHandler.Animations.cs
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput, notFound } from "./result.js";
import type { Result, AnimationModel } from "./types.js";
import { getSlideIndex } from "./path.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Animation trigger types.
 */
export type AnimationTrigger = "onClick" | "afterPrev" | "withPrev" | "onLoad";

/**
 * Animation class/types.
 */
export type AnimationClass = "entrance" | "exit" | "emphasis" | "motionPath";

/**
 * Specification for adding an animation.
 */
export interface AnimationSpec {
  /** Animation effect name (e.g., "fade", "fly", "zoom", "appear") */
  effect: string;
  /** Animation class: entrance, exit, emphasis (default: entrance) */
  class?: AnimationClass;
  /** Animation trigger: onClick, afterPrev, withPrev, onLoad (default: onClick) */
  trigger?: AnimationTrigger;
  /** Duration in milliseconds (default: 500) */
  duration?: number;
  /** Delay in milliseconds (default: 0) */
  delay?: number;
  /** Direction for directional effects: left, right, up, down */
  direction?: "left" | "right" | "up" | "down";
  /** Ease-in percentage (0-100) */
  easein?: number;
  /** Ease-out percentage (0-100) */
  easeout?: number;
}

/**
 * Result from getting animations on a slide.
 */
export interface SlideAnimationsResult {
  /** Slide index (1-based) */
  slideIndex: number;
  /** Slide path */
  path: string;
  /** All animations on the slide */
  animations: AnimationModel[];
  /** Count of animations */
  count: number;
}

// ============================================================================
// Animation Preset Definitions (from OOXML specification)
// ============================================================================

/**
 * Animation effect presets mapping effect names to (presetId, filter) tuples.
 */
const ANIM_PRESETS: Record<string, { entrance: [number, string | null]; exit: [number, string | null]; emphasis: [number, string | null] }> = {
  appear:    { entrance: [1, null],       exit: [1, null],       emphasis: [1, null] },
  fly:       { entrance: [2, null],       exit: [2, null],       emphasis: [1, null] },
  flyin:     { entrance: [2, null],       exit: [2, null],       emphasis: [1, null] },
  flyout:    { entrance: [2, null],       exit: [2, null],       emphasis: [1, null] },
  blinds:    { entrance: [3, "blinds(horizontal)"], exit: [3, "blinds(horizontal)"], emphasis: [1, null] },
  box:       { entrance: [4, "box"],      exit: [4, "box"],      emphasis: [1, null] },
  checkerboard: { entrance: [5, "checkerboard(across)"], exit: [5, "checkerboard(across)"], emphasis: [1, null] },
  circle:    { entrance: [6, "circle"],   exit: [6, "circle"],   emphasis: [1, null] },
  crawl:     { entrance: [7, "crawl"],    exit: [7, "crawl"],    emphasis: [1, null] },
  diamond:   { entrance: [8, "diamond"], exit: [8, "diamond"], emphasis: [1, null] },
  dissolve:  { entrance: [9, "dissolve"], exit: [9, "dissolve"], emphasis: [1, null] },
  fade:      { entrance: [10, "fade"],   exit: [10, "fade"],    emphasis: [10, "fade"] },
  flash:     { entrance: [11, "flash"],  exit: [11, "flash"],   emphasis: [1, null] },
  float:     { entrance: [12, null],     exit: [12, null],      emphasis: [1, null] },
  plus:      { entrance: [13, "plus"],   exit: [13, "plus"],    emphasis: [1, null] },
  random:    { entrance: [14, "random"], exit: [14, "random"], emphasis: [14, null] },
  split:     { entrance: [15, "barn(inHorizontal)"], exit: [15, "barn(inHorizontal)"], emphasis: [1, null] },
  strips:    { entrance: [16, "strips(downLeft)"], exit: [16, "strips(downLeft)"], emphasis: [1, null] },
  swivel:    { entrance: [17, null],      exit: [17, null],      emphasis: [1, null] },
  wedge:     { entrance: [18, "wedge"],   exit: [18, "wedge"],   emphasis: [1, null] },
  wheel:     { entrance: [19, "wheel(1)"], exit: [19, "wheel(1)"], emphasis: [1, null] },
  wipe:      { entrance: [20, "wipe(left)"], exit: [20, "wipe(left)"], emphasis: [1, null] },
  zoom:      { entrance: [21, null],      exit: [21, null],      emphasis: [1, null] },
  bounce:    { entrance: [24, null],     exit: [24, null],      emphasis: [1, null] },
  bold:      { entrance: [1, null],      exit: [1, null],       emphasis: [1, null] },
  grow:      { entrance: [1, null],      exit: [1, null],       emphasis: [26, null] },
  shrink:    { entrance: [1, null],      exit: [1, null],       emphasis: [26, null] },
  wave:      { entrance: [1, null],      exit: [1, null],       emphasis: [14, null] },
  spin:      { entrance: [1, null],      exit: [1, null],       emphasis: [27, null] },
  rotate:    { entrance: [1, null],      exit: [1, null],       emphasis: [27, null] },
};

/**
 * Get preset ID and filter for an animation effect.
 */
function getAnimPreset(effect: string, animClass: AnimationClass): { presetId: number; filter: string | null } {
  const effectLower = effect.toLowerCase();
  const preset = ANIM_PRESETS[effectLower];

  if (!preset) {
    const validEffects = Object.keys(ANIM_PRESETS).filter(e => !["bold", "grow", "shrink", "wave", "spin", "rotate"].includes(e) || animClass === "emphasis");
    throw new Error(`Unknown animation effect: '${effect}'. Valid effects: ${validEffects.join(", ")}`);
  }

  switch (animClass) {
    case "exit":
      return { presetId: preset.exit[0], filter: preset.exit[1] };
    case "emphasis":
      return { presetId: preset.emphasis[0], filter: preset.emphasis[1] };
    default:
      return { presetId: preset.entrance[0], filter: preset.entrance[1] };
  }
}

/**
 * Get preset subtype based on direction.
 * Subtypes: 0=none, 1=from-top, 2=from-right, 4=from-bottom, 8=from-left
 */
function getAnimPresetSubtype(effect: string, direction: string | undefined): number {
  if (direction) {
    switch (direction) {
      case "left":  return 8;  // from left
      case "right": return 2;  // from right
      case "up":    return 1;  // from top
      case "down":  return 4;  // from bottom
    }
  }

  // Effect-specific defaults
  switch (effect.toLowerCase()) {
    case "fly":
    case "flyin":
    case "flyout":
      return 4;  // from bottom
    case "wipe":
      return 1;  // from left
    case "blinds":
      return 10; // horizontal
    case "checkerboard":
    case "checker":
      return 5;  // across
    case "strips":
      return 7;  // down-left
    case "split":
      return 10; // horizontal in
    case "wheel":
      return 1; // 1 spoke
    default:
      return 0;
  }
}

// ============================================================================
// Helpers
// ============================================================================

/**
 * Parses relationship entries from a .rels XML string.
 */
function parseRelationshipEntries(xml: string): Array<{ id: string; target: string; type?: string }> {
  const relationships: Array<{ id: string; target: string; type?: string }> = [];
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const attributes = match[1];
    const id = /Id="([^"]+)"/.exec(attributes)?.[1];
    const target = /Target="([^"]+)"/.exec(attributes)?.[1];
    const type = /Type="([^"]+)"/.exec(attributes)?.[1];
    if (id && target) {
      relationships.push({ id, target, type });
    }
  }
  return relationships;
}

/**
 * Normalizes a zip path relative to a base directory.
 */
function normalizeZipPath(baseDir: string, target: string): string {
  const normalized = target.replace(/\\/g, "/");
  if (normalized.startsWith("/")) {
    return path.posix.normalize(normalized.slice(1));
  }
  return path.posix.normalize(path.posix.join(baseDir, normalized));
}

/**
 * Reads an entry from the zip as a string.
 */
function requireEntry(zip: Map<string, Buffer>, entryName: string): string {
  const buffer = zip.get(entryName);
  if (!buffer) {
    throw new Error(`OOXML entry '${entryName}' is missing`);
  }
  return buffer.toString("utf8");
}

/**
 * Gets the slide IDs from presentation.xml.
 */
function getSlideIds(presentationXml: string): Array<{ id: string; relId: string }> {
  const slideIds: Array<{ id: string; relId: string }> = [];
  for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*\bid="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g)) {
    slideIds.push({ id: match[1], relId: match[2] });
  }
  for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"[^>]*\bid="([^"]+)"[^>]*\/?>/g)) {
    const relId = match[1];
    const id = match[2];
    if (!slideIds.some(s => s.relId === relId)) {
      slideIds.push({ id, relId });
    }
  }
  return slideIds;
}

/**
 * Gets the slide entry path from the zip by slide index.
 */
function getSlideEntryPath(zip: Map<string, Buffer>, slideIndex: number): Result<string> {
  const presentationXml = requireEntry(zip, "ppt/presentation.xml");
  const relsXml = requireEntry(zip, "ppt/_rels/presentation.xml.rels");
  const relationships = parseRelationshipEntries(relsXml);
  const slideIds = getSlideIds(presentationXml);

  if (slideIndex < 1 || slideIndex > slideIds.length) {
    return invalidInput(`Slide index ${slideIndex} is out of range (1-${slideIds.length})`);
  }

  const slide = slideIds[slideIndex - 1];
  const slideRel = relationships.find(r => r.id === slide.relId);
  const slidePath = normalizeZipPath("ppt", slideRel?.target ?? "");

  return ok(slidePath);
}

/**
 * Loads a presentation and returns its zip contents.
 */
async function loadPresentation(filePath: string): Promise<Result<Map<string, Buffer>>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);
    return ok(zip);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Escapes special XML characters.
 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// ============================================================================
// Animation XML Building
// ============================================================================

/**
 * Generates unique IDs for timing nodes.
 */
let nextTimingId = 1;
let nextGroupId = 0;

function getNextTimingId(): number {
  return nextTimingId++;
}

function getNextGroupId(): number {
  return nextGroupId++;
}

function resetIdCounters(timingXml: string): void {
  // Find max existing IDs in the timing XML
  const idMatches = timingXml.match(/id="(\d+)"/g) || [];
  let maxId = 0;
  for (const match of idMatches) {
    const id = parseInt(match.match(/id="(\d+)"/)![1], 10);
    if (id > maxId) maxId = id;
  }
  nextTimingId = maxId + 1;

  // Find max group IDs
  const grpMatches = timingXml.match(/grpId="(\d+)"/g) || [];
  let maxGrp = -1;
  for (const match of grpMatches) {
    const grpId = parseInt(match.match(/grpId="(\d+)"/)![1], 10);
    if (grpId > maxGrp) maxGrp = grpId;
  }
  nextGroupId = maxGrp + 1;
}

/**
 * Build animation XML for a shape based on the animation spec.
 * Returns the XML string for a p:par element containing the animation.
 */
function buildAnimationXml(
  shapeId: string,
  spec: AnimationSpec,
  existingTimingXml: string | null,
  triggerType: string
): string {
  // Reset and initialize ID counters
  nextTimingId = 1;
  nextGroupId = 0;
  if (existingTimingXml) {
    resetIdCounters(existingTimingXml);
  }

  const animClass = spec.class || "entrance";
  const duration = spec.duration || 500;
  const delay = spec.delay || 0;
  const direction = spec.direction;
  const easein = spec.easein || 0;
  const easeout = spec.easeout || 0;

  // Get preset info
  const { presetId, filter } = getAnimPreset(spec.effect, animClass);
  const presetSubtype = getAnimPresetSubtype(spec.effect, direction);

  // Map animation class to OOXML preset class
  const presetClass = animClass === "exit" ? "exit" : animClass === "emphasis" ? "emph" : "entr";
  const isEntrance = animClass === "entrance";
  const isEmphasis = animClass === "emphasis";

  // Determine node type based on trigger
  const nodeType = triggerType === "afterPrev" ? "afterEffect"
    : triggerType === "withPrev" ? "withEffect"
    : triggerType === "onLoad" ? "onLoad"
    : "clickEffect";

  // Outer delay for click trigger
  const outerDelay = triggerType === "onClick" ? "indefinite" : "0";

  // Get inner animation type
  const innerAnimType = presetId === 2 || presetId === 12 ? "fly"
    : presetId === 21 ? "zoom"
    : presetId === 17 ? "swivel"
    : filter ? "standard"
    : "none";

  const innerDuration = innerAnimType !== "none" && innerAnimType !== "fly" ? duration : 0;

  // Build the animation par structure
  // Structure: p:par > outer cTn > mid cTn > effect cTn > behaviors

  const effectId = getNextTimingId();
  const animVisId = getNextTimingId();
  const animEffId = innerAnimType !== "none" ? getNextTimingId() : 0;

  // Build set behavior for visibility
  let setBehaviorXml = "";
  if (isEntrance || isEmphasis) {
    setBehaviorXml = `<p:set>
  <p:cBhvr>
    <p:cTn id="${animVisId}" dur="1" fill="hold">
      <p:stCondLst>
        <p:cond delay="0"/>
      </p:stCondLst>
    </p:cTn>
    <p:tavLst>
      <p:ta>
        <p:stAttr name="style.visibility" display="0"/>
        <p:to>
          <p:strVal val="visible"/>
        </p:to>
      </p:ta>
    </p:tavLst>
  </p:cBhvr>
  <p:to>
    <p:strVal val="visible"/>
  </p:to>
</p:set>`;
  } else {
    // Exit: make hidden
    setBehaviorXml = `<p:set>
  <p:cBhvr>
    <p:cTn id="${animVisId}" dur="1" fill="hold">
      <p:stCondLst>
        <p:cond delay="0"/>
      </p:stCondLst>
    </p:cTn>
    <p:tavLst>
      <p:ta>
        <p:stAttr name="style.visibility" display="0"/>
        <p:to>
          <p:strVal val="hidden"/>
        </p:to>
      </p:ta>
    </p:tavLst>
  </p:cBhvr>
  <p:to>
    <p:strVal val="hidden"/>
  </p:to>
</p:set>`;
  }

  // Build effect-specific animation element
  let effectAnimXml = "";
  if (innerAnimType === "fly") {
    // Fly animation uses ppt_x or ppt_y property animation
    const axis = presetSubtype === 8 ? "ppt_x" : presetSubtype === 2 ? "ppt_x" : "ppt_y";
    const startVal = presetSubtype === 8 ? "0-#ppt_w/2"  // from left
      : presetSubtype === 2 ? "1+#ppt_w/2"  // from right
      : presetSubtype === 1 ? "0-#ppt_h/2"  // from top
      : "1+#ppt_h/2"; // from bottom (default)
    const endVal = "#ppt_x";

    effectAnimXml = `<p:anim calcmode="lin">
  <p:cBhvr additive="base">
    <p:cTn id="${animEffId}" dur="${duration}" fill="hold">
      <p:stCondLst>
        <p:cond delay="0"/>
      </p:stCondLst>
    </p:cTn>
    <p:tavLst>
      <p:ta>
        <p:stAttr name="${axis}"/>
        <p:to>
          <p:strVal val="${endVal}"/>
        </p:to>
      </p:ta>
    </p:tavLst>
    <p:target>
      <p:spTgt spid="${shapeId}"/>
    </p:target>
  </p:cBhvr>
</p:anim>`;
  } else if (innerAnimType === "zoom") {
    // Zoom animation uses animScale
    effectAnimXml = `<p:animScale>
  <p:cBhvr>
    <p:cTn id="${animEffId}" dur="${duration}" fill="hold">
      <p:stCondLst>
        <p:cond delay="0"/>
      </p:stCondLst>
    </p:cTn>
    <p:target>
      <p:spTgt spid="${shapeId}"/>
    </p:target>
  </p:cBhvr>
  <p:to x="100000" y="100000"/>
</p:animScale>`;
  } else if (innerAnimType === "swivel") {
    // Swivel = rotation + fade
    effectAnimXml = `<p:animRot by="21600000">
  <p:cBhvr>
    <p:cTn id="${animEffId}" dur="${duration}" fill="hold">
      <p:stCondLst>
        <p:cond delay="0"/>
      </p:stCondLst>
    </p:cTn>
    <p:target>
      <p:spTgt spid="${shapeId}"/>
    </p:target>
  </p:cBhvr>
</p:animRot>
<p:animEffect transition="in" filter="fade">
  <p:cBhvr>
    <p:cTn id="${getNextTimingId()}" dur="${duration}">
      <p:stCondLst>
        <p:cond delay="0"/>
      </p:stCondLst>
    </p:cTn>
    <p:target>
      <p:spTgt spid="${shapeId}"/>
    </p:target>
  </p:cBhvr>
</p:animEffect>`;
  } else if (filter) {
    // Standard animEffect
    const transition = isEntrance || isEmphasis ? "in" : "out";
    effectAnimXml = `<p:animEffect transition="${transition}" filter="${escapeXml(filter)}">
  <p:cBhvr>
    <p:cTn id="${animEffId}" dur="${duration}" fill="hold">
      <p:stCondLst>
        <p:cond delay="0"/>
      </p:stCondLst>
    </p:cTn>
    <p:target>
      <p:spTgt spid="${shapeId}"/>
    </p:target>
  </p:cBhvr>
</p:animEffect>`;
  }

  // Build the effect cTn (innermost)
  let effectCtnXml = `p:cTn id="${effectId}" presetId="${presetId}" presetClass="${presetClass}" presetSubtype="${presetSubtype}" fill="hold" nodeType="${nodeType}">`;
  if (innerAnimType === "none") {
    effectCtnXml += ` dur="${duration}"`;
  }
  if (easein > 0) {
    effectCtnXml += ` accel="${easein * 100}"`;
  }
  if (easeout > 0) {
    effectCtnXml += ` decel="${easeout * 100}"`;
  }
  effectCtnXml += `>
    <p:stCondLst>
      <p:cond delay="0"/>
    </p:stCondLst>
    <p:childTnLst>
      ${setBehaviorXml}
      ${effectAnimXml}
    </p:childTnLst>
  </p:cTn>`;

  const midId = getNextTimingId();
  const midCtnXml = `p:cTn id="${midId}" fill="hold">
    <p:stCondLst>
      <p:cond delay="${delay}"/>
    </p:stCondLst>
    <p:childTnLst>
      <p:par>${effectCtnXml}</p:par>
    </p:childTnLst>
  </p:cTn>`;

  const outerId = getNextTimingId();
  const groupId = getNextGroupId();
  const outerCtnXml = `p:cTn id="${outerId}" fill="hold" nodeType="clickGroup">
    <p:stCondLst>
      <p:cond delay="${outerDelay}"/>
    </p:stCondLst>
    <p:childTnLst>
      <p:par>${midCtnXml}</p:par>
    </p:childTnLst>
  </p:cTn>`;

  return `<p:par>${outerCtnXml}</p:par>`;
}

/**
 * Ensures a timing tree exists in the slide XML and returns the updated XML.
 * Creates the basic timing structure if it doesn't exist.
 */
function ensureTimingTree(slideXml: string): { slideXml: string; hasExistingAnimations: boolean } {
  const timingMatch = slideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/);
  let hasExistingAnimations = false;

  if (timingMatch) {
    // Check if there are any existing animations
    const timingXml = timingMatch[0];
    hasExistingAnimations = timingXml.includes('presetId="') || timingXml.includes("p:anim");
  }

  if (!timingMatch) {
    // Create new timing element
    const newTiming = `<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAction="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:prevCondLst>
                <p:cond evt="onPrev" delay="0">
                  <p:tgtEl>
                    <p:sldTgt/>
                  </p:tgtEl>
                </p:cond>
              </p:prevCondLst>
              <p:nextCondLst>
                <p:cond evt="onNext" delay="0">
                  <p:tgtEl>
                    <p:sldTgt/>
                  </p:tgtEl>
                </p:cond>
              </p:nextCondLst>
            </p:cTn>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>`;

    // Insert before </p:sld> or </p:spTree>
    if (slideXml.includes("</p:spTree>")) {
      slideXml = slideXml.replace("</p:spTree>", `${newTiming}
</p:spTree>`);
    } else if (slideXml.includes("</p:sld>")) {
      slideXml = slideXml.replace("</p:sld>", `${newTiming}
</p:sld>`);
    }
  }

  return { slideXml, hasExistingAnimations };
}

/**
 * Gets the main sequence (mainSeq) cTn ID from timing XML.
 */
function getMainSeqId(timingXml: string): number {
  const match = timingXml.match(/p:cTn id="(\d+)"[^>]*nodeType="mainSeq"/);
  return match ? parseInt(match[1], 10) : 2;
}

/**
 * Inserts animation XML into the main sequence of timing.
 */
function insertAnimationIntoTiming(timingXml: string, animXml: string): string {
  // Find the mainSeq cTn and insert after its opening
  const mainSeqMatch = timingXml.match(/(<p:cTn id="\d+"[^>]*nodeType="mainSeq"[^>]*>[\s\S]*?<p:childTnLst>)/);
  if (mainSeqMatch && mainSeqMatch.index !== undefined) {
    const insertPoint = mainSeqMatch[0].length - "<p:childTnLst>".length;
    const before = timingXml.substring(0, mainSeqMatch.index + insertPoint);
    const after = timingXml.substring(mainSeqMatch.index + insertPoint);
    return before + after + `
      ${animXml}
    `;
  }

  // Fallback: insert before </p:tnLst>
  return timingXml.replace("</p:tnLst>", `${animXml}
    </p:tnLst>`);
}

// ============================================================================
// Animation Extraction (Reading)
// ============================================================================

/**
 * Extracts animations from a slide's timing XML.
 */
function extractAnimationsFromTiming(timingXml: string, slideIndex: number): AnimationModel[] {
  const animations: AnimationModel[] = [];

  // Find all p:par elements in timing
  const parPattern = /<p:par>[\s\S]*?<\/p:par>/g;
  const parMatches = timingXml.match(parPattern) || [];

  for (const parXml of parMatches) {
    // Check if this is a click group with an animation
    const clickGroupMatch = parXml.match(/<p:cTn[^>]*nodeType="clickGroup"[^>]*>[\s\S]*?<\/p:cTn>/);
    if (!clickGroupMatch) continue;

    const clickGroupXml = clickGroupMatch[0];

    // Get preset ID and class
    const presetIdMatch = clickGroupXml.match(/presetId="(\d+)"/);
    const presetClassMatch = clickGroupXml.match(/presetClass="([^"]+)"/);
    const presetSubtypeMatch = clickGroupXml.match(/presetSubtype="(\d+)"/);

    if (!presetIdMatch) continue;

    const presetId = parseInt(presetIdMatch[1], 10);
    const presetClassRaw = presetClassMatch ? presetClassMatch[1] : "";

    // Map preset class to animation class
    let animClass: string | undefined;
    if (presetClassRaw === "entr") {
      animClass = "entrance";
    } else if (presetClassRaw === "exit") {
      animClass = "exit";
    } else if (presetClassRaw === "emph") {
      animClass = "emphasis";
    } else if (presetClassRaw === "motion") {
      animClass = "motionPath";
    }

    // Get effect name from preset ID
    const effectName = getEffectNameFromPresetId(presetId, animClass || "entrance");

    // Get duration
    let duration: number | undefined;
    const durMatch = clickGroupXml.match(/dur="(\d+)"/);
    if (durMatch) {
      duration = parseInt(durMatch[1], 10);
    }

    // Get delay
    let delay: number | undefined;
    const delayMatch = parXml.match(/<p:cond[^>]*delay="([^"]*)"[^>]*>/);
    if (delayMatch) {
      const delayStr = delayMatch[1];
      if (delayStr && delayStr !== "0" && delayStr !== "indefinite") {
        delay = parseInt(delayStr, 10);
      }
    }

    // Get direction from preset subtype
    let direction: string | undefined;
    if (presetSubtypeMatch) {
      const subtype = parseInt(presetSubtypeMatch[1], 10);
      direction = subtypeToDirection(subtype, effectName);
    }

    // Get trigger type
    const nodeTypeMatch = clickGroupXml.match(/nodeType="([^"]*)"/);
    let trigger: AnimationTrigger | undefined;
    if (nodeTypeMatch) {
      const nodeType = nodeTypeMatch[1];
      if (nodeType === "clickEffect") {
        trigger = "onClick";
      } else if (nodeType === "afterEffect") {
        trigger = "afterPrev";
      } else if (nodeType === "withEffect") {
        trigger = "withPrev";
      } else if (nodeType === "onLoad") {
        trigger = "onLoad";
      }
    }

    // Find the target shape ID
    const targetMatch = parXml.match(/<p:spTgt[^>]*spid="([^"]*)"[^>]*>/);
    let path: string;
    if (targetMatch) {
      path = `/slide[${slideIndex}]/shape[${targetMatch[1]}]`;
    } else {
      path = `/slide[${slideIndex}]`;
    }

    animations.push({
      path,
      effect: effectName,
      class: animClass,
      presetId,
      duration,
      delay,
      ...(direction && { direction: direction as any }),
    });
  }

  return animations;
}

/**
 * Maps preset ID to effect name.
 */
function getEffectNameFromPresetId(presetId: number, animClass: string): string {
  const entranceNames: Record<number, string> = {
    1: "appear", 2: "fly", 3: "blinds", 4: "box", 5: "checkerboard",
    6: "circle", 7: "crawl", 8: "diamond", 9: "dissolve", 10: "fade",
    11: "flash", 12: "float", 13: "plus", 14: "random", 15: "split",
    16: "strips", 17: "swivel", 18: "wedge", 19: "wheel", 20: "wipe",
    21: "zoom", 24: "bounce"
  };

  if (animClass === "exit") {
    return entranceNames[presetId] || "unknown";
  }

  if (animClass === "emphasis") {
    switch (presetId) {
      case 1: return "bold";
      case 10: return "fade";
      case 14: return "wave";
      case 26: return "grow";
      case 27: return "spin";
      default: return entranceNames[presetId] || "unknown";
    }
  }

  return entranceNames[presetId] || "unknown";
}

/**
 * Maps preset subtype to direction string.
 */
function subtypeToDirection(subtype: number, effectName: string): string | undefined {
  switch (subtype) {
    case 8: return "left";
    case 2: return "right";
    case 1: return effectName === "fly" || effectName === "wipe" || effectName === "crawl" ? "up" : "top";
    case 4: return effectName === "fly" || effectName === "wipe" || effectName === "crawl" ? "down" : "bottom";
    default: return undefined;
  }
}

/**
 * Gets animations from a slide's timing XML.
 */
function getAnimationsFromSlideXml(slideXml: string, slideIndex: number): AnimationModel[] {
  const timingMatch = slideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/);
  if (!timingMatch) {
    return [];
  }
  return extractAnimationsFromTiming(timingMatch[0], slideIndex);
}

// ============================================================================
// Animation Removal
// ============================================================================

/**
 * Removes all animations for a specific shape from timing XML.
 */
function removeAnimationsForShape(timingXml: string, shapeId: string): string {
  // Find all p:par elements that target this shape
  const parPattern = /<p:par>[\s\S]*?<\/p:par>/g;
  let result = timingXml;
  let match;

  const toRemove: string[] = [];

  while ((match = parPattern.exec(timingXml)) !== null) {
    const parXml = match[0];
    // Check if this par targets the shape
    const targetMatch = parXml.match(/<p:spTgt[^>]*spid="([^"]*)"[^>]*>/);
    if (targetMatch && targetMatch[1] === shapeId) {
      // Check if this is a clickGroup animation
      if (parXml.includes('nodeType="clickGroup"')) {
        toRemove.push(parXml);
      }
    }
  }

  for (const xml of toRemove) {
    result = result.replace(xml, "");
  }

  // Clean up empty timing
  result = result.replace(/\s*<p:par>\s*<\/p:par>\s*/g, "");

  // If timing is empty, remove it entirely
  if (result.match(/<p:timing>[\s]*<\/p:timing>/)) {
    result = result.replace(/<p:timing>[\s]*<\/p:timing>/, "");
  }

  return result;
}

// ============================================================================
// Public API
// ============================================================================

/**
 * Gets all animations on a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result with all animations on the slide
 *
 * @example
 * const result = await getAnimations("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(`Found ${result.data.count} animations`);
 *   for (const anim of result.data.animations) {
 *     console.log(`  ${anim.effect} (${anim.class})`);
 *   }
 * }
 */
export async function getAnimations(
  filePath: string,
  slideIndex: number,
): Promise<Result<SlideAnimationsResult>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code ?? "load_failed", zipResult.error?.message ?? "Failed to load presentation");
  }
  const zip = zipResult.data;
  if (!zip) {
    return err("operation_failed", "Failed to load presentation");
  }

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
  }

  const slideEntry = slidePathResult.data;
  if (!slideEntry) {
    return err("slide_not_found", "Slide entry not found");
  }
  const slideXml = requireEntry(zip, slideEntry);

  const animations = getAnimationsFromSlideXml(slideXml, slideIndex);

  return ok({
    slideIndex,
    path: `/slide[${slideIndex}]`,
    animations,
    count: animations.length,
  });
}

/**
 * Adds an animation to a shape.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @param animation - Animation specification
 * @returns Result indicating success
 *
 * @example
 * // Add a fade entrance animation
 * const result = await addAnimation("/path/to/presentation.pptx", "/slide[1]/shape[1]", {
 *   effect: "fade",
 *   class: "entrance",
 *   trigger: "onClick",
 *   duration: 500
 * });
 *
 * // Add a fly emphasis animation from the left
 * const result = await addAnimation("/path/to/presentation.pptx", "/slide[1]/shape[2]", {
 *   effect: "fly",
 *   class: "entrance",
 *   direction: "left",
 *   duration: 400,
 *   delay: 1000
 * });
 *
 * // Add a zoom exit animation
 * const result = await addAnimation("/path/to/presentation.pptx", "/slide[1]/shape[3]", {
 *   effect: "zoom",
 *   class: "exit",
 *   trigger: "afterPrev",
 *   duration: 300
 * });
 */
export async function addAnimation(
  filePath: string,
  pptPath: string,
  animation: AnimationSpec,
): Promise<Result<void>> {
  const slideIndex = getSlideIndex(pptPath);
  if (slideIndex === null) {
    return invalidInput("addAnimation requires a slide path (e.g., /slide[1]/shape[1])");
  }

  // Extract shape index from path
  const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
  if (!shapeIndexMatch) {
    return invalidInput("Invalid shape path. Expected format: /slide[N]/shape[N]");
  }
  const shapeId = shapeIndexMatch[1];

  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);

    // First, remove any existing animations on this shape
    let updatedSlideXml = removeAnimationsForShapeFromSlideXml(slideXml, shapeId);

    // Ensure timing tree exists
    const { slideXml: withTiming, hasExistingAnimations } = ensureTimingTree(updatedSlideXml);
    updatedSlideXml = withTiming;

    // Determine trigger type
    let triggerType: string = "click";
    if (animation.trigger === "afterPrev") {
      triggerType = "afterPrev";
    } else if (animation.trigger === "withPrev") {
      triggerType = "withPrev";
    } else if (animation.trigger === "onLoad") {
      triggerType = "onLoad";
    } else if (animation.trigger === "onClick") {
      triggerType = "onClick";
    } else if (!hasExistingAnimations) {
      // First animation defaults to onClick
      triggerType = "onClick";
    } else {
      // Subsequent animations default to afterPrev
      triggerType = "afterPrev";
    }

    // Get existing timing XML for ID counter initialization
    const existingTimingMatch = updatedSlideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/);
    const existingTimingXml = existingTimingMatch ? existingTimingMatch[0] : null;

    // Build animation XML
    const animXml = buildAnimationXml(shapeId, animation, existingTimingXml, triggerType);

    // Insert into timing
    const timingMatch = updatedSlideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/);
    if (timingMatch) {
      const updatedTiming = insertAnimationIntoTiming(timingMatch[0], animXml);
      updatedSlideXml = updatedSlideXml.replace(timingMatch[0], updatedTiming);
    }

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));
    return ok(void 0);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Removes animations for a specific shape from slide XML.
 */
function removeAnimationsForShapeFromSlideXml(slideXml: string, shapeId: string): string {
  const timingMatch = slideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/);
  if (!timingMatch) {
    return slideXml; // No timing element, nothing to remove
  }

  const timingXml = timingMatch[0];
  const timingIndex = slideXml.indexOf(timingXml);

  const updatedTiming = removeAnimationsForShape(timingXml, shapeId);

  // If timing is now empty, remove it
  if (updatedTiming.match(/<p:timing>[\s]*<\/p:timing>/) || updatedTiming.trim() === "<p:timing>") {
    return slideXml.slice(0, timingIndex) + slideXml.slice(timingIndex + timingXml.length);
  }

  return slideXml.replace(timingXml, updatedTiming);
}

/**
 * Removes an animation from an element.
 *
 * @param filePath - Path to the PPTX file
 * @param pptPath - PPT path to the shape (e.g., "/slide[1]/shape[1]")
 * @returns Result indicating success
 *
 * @example
 * const result = await removeAnimation("/path/to/presentation.pptx", "/slide[1]/shape[1]");
 */
export async function removeAnimation(
  filePath: string,
  pptPath: string,
): Promise<Result<void>> {
  const slideIndex = getSlideIndex(pptPath);
  if (slideIndex === null) {
    return invalidInput("removeAnimation requires a slide path (e.g., /slide[1]/shape[1])");
  }

  // Extract shape index from path
  const shapeIndexMatch = pptPath.match(/\/shape\[(\d+)\]/i);
  if (!shapeIndexMatch) {
    return invalidInput("Invalid shape path. Expected format: /slide[N]/shape[N]");
  }
  const shapeId = shapeIndexMatch[1];

  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);

    // Remove animations for this shape
    const updatedSlideXml = removeAnimationsForShapeFromSlideXml(slideXml, shapeId);

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));
    return ok(void 0);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// Morph Transition (Office 365 / PowerPoint 2016+)
// ============================================================================

/**
 * Morph transition types.
 */
export type MorphType = "object" | "text" | "chart" | "table" | "picture";

/**
 * Morph effect options.
 */
export interface MorphSpec {
  /** Morph type: object, text, chart, table, picture (default: object) */
  type?: MorphType;
  /** Duration in milliseconds (default: 1000) */
  duration?: number;
  /** Ease-in percentage (0-100) */
  easein?: number;
  /** Ease-out percentage (0-100) */
  easeout?: number;
}

/**
 * Detected morph node information.
 */
export interface MorphNode {
  /** Shape ID of the morph node */
  shapeId: string;
  /** Shape name */
  name?: string;
  /** Morph type: "object", "text", etc. */
  type: string;
  /** Preset ID for the morph */
  presetId: number;
}

/**
 * Result from detecting morph nodes on a slide.
 */
export interface MorphNodesResult {
  /** Slide index (1-based) */
  slideIndex: number;
  /** Slide path */
  path: string;
  /** Detected morph nodes */
  nodes: MorphNode[];
  /** Count of morph nodes */
  count: number;
}

/**
 * Gets unique IDs for morph animation elements.
 */
let nextMorphId = 1;

function getNextMorphId(): number {
  return nextMorphId++;
}

function resetMorphIdCounter(timingXml: string): void {
  const idMatches = timingXml.match(/id="(\d+)"/g) || [];
  let maxId = 0;
  for (const match of idMatches) {
    const id = parseInt(match.match(/id="(\d+)"/)![1], 10);
    if (id > maxId) maxId = id;
  }
  nextMorphId = maxId + 1;
}

/**
 * Detects morph nodes on a slide that can be used for Morph transitions.
 * Morph nodes are shapes with specific animation properties that enable
 * the Morph transition effect between consecutive slides.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result with detected morph nodes
 *
 * @example
 * const result = await detectMorphNodes("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(`Found ${result.data.count} morph nodes`);
 *   for (const node of result.data.nodes) {
 *     console.log(`  ${node.name}: ${node.type}`);
 *   }
 * }
 */
export async function detectMorphNodes(
  filePath: string,
  slideIndex: number,
): Promise<Result<MorphNodesResult>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code ?? "load_failed", zipResult.error?.message ?? "Failed to load presentation");
  }
  const zip = zipResult.data;
  if (!zip) {
    return err("operation_failed", "Failed to load presentation");
  }

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
  }

  const slideEntry = slidePathResult.data;
  if (!slideEntry) {
    return err("slide_not_found", "Slide entry not found");
  }
  const slideXml = requireEntry(zip, slideEntry);

  // Find all shapes with morph-compatible animations
  const nodes: MorphNode[] = [];

  // Look for animBg elements (indicates morph background)
  const animBgMatches = slideXml.match(/<p:animBg\b[^>]*\/>/g) || [];
  for (const match of animBgMatches) {
    const presetIdMatch = match.match(/presetID="(\d+)"/i);
    const shapeIdMatch = match.match(/spid="([^"]+)"/i);
    const nameMatch = match.match(/name="([^"]+)"/i);
    nodes.push({
      shapeId: shapeIdMatch?.[1] ?? "bg",
      name: nameMatch?.[1] ?? "Background",
      type: "object",
      presetId: presetIdMatch ? parseInt(presetIdMatch[1], 10) : 0,
    });
  }

  // Look for p:pt (morph point) elements in timing
  const timingMatch = slideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/);
  if (timingMatch) {
    const timingXml = timingMatch[0];

    // Find morph effect elements
    const morphPattern = /<p:morph\b[^>]*\/?>/g;
    const morphMatches = timingXml.match(morphPattern) || [];

    for (const morphXml of morphMatches) {
      const shapeIdMatch = morphXml.match(/spid="([^"]+)"/i);
      const presetIdMatch = morphXml.match(/presetID="(\d+)"/i);
      const typeMatch = morphXml.match(/morphType="([^"]+)"/i);

      if (shapeIdMatch) {
        // Get shape name from slide XML
        const shapeId = shapeIdMatch[1];
        const shapeNameMatch = slideXml.match(new RegExp(`<p:sp\b[^>]*id="${shapeId}"[^>]*name="([^"]*)"`, "i"));
        nodes.push({
          shapeId,
          name: shapeNameMatch?.[1],
          type: typeMatch?.[1] ?? "object",
          presetId: presetIdMatch ? parseInt(presetIdMatch[1], 10) : 0,
        });
      }
    }
  }

  return ok({
    slideIndex,
    path: `/slide[${slideIndex}]`,
    nodes,
    count: nodes.length,
  });
}

/**
 * Creates morph effect XML for a shape transition.
 * Morph allows smooth transitions between slides where elements
 * appear to transform, move, or change in size/position.
 *
 * @param shapeId - Shape ID to apply morph to
 * @param spec - Morph specification
 * @returns XML string for the morph effect
 */
function buildMorphEffectXml(shapeId: string, spec: MorphSpec): string {
  const morphType = spec.type ?? "object";
  const duration = spec.duration ?? 1000;
  const easein = spec.easein ?? 0;
  const easeout = spec.easeout ?? 0;

  // Map morph type to OOXML morphType attribute
  const morphTypeMap: Record<string, string> = {
    object: "obj",
    text: "txt",
    chart: "chart",
    table: "tbl",
    picture: "pict",
  };
  const ooxmlMorphType = morphTypeMap[morphType] ?? "obj";

  const morphId = getNextMorphId();

  // Build morph XML structure
  let morphXml = `<p:morph morphType="${ooxmlMorphType}" spid="${shapeId}" presetID="${morphId}"`;
  if (easein > 0) {
    morphXml += ` accel="${easein * 100}"`;
  }
  if (easeout > 0) {
    morphXml += ` decel="${easeout * 100}"`;
  }
  morphXml += `><p:cBhvr><p:cTn id="${morphId + 1}" dur="${duration}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tavLst></p:tavLst></p:cBhvr></p:morph>`;

  return morphXml;
}

/**
 * Checks if a slide has morph transition configured.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result indicating if morph is configured
 *
 * @example
 * const result = await hasMorphTransition("/path/to/presentation.pptx", 1);
 * if (result.ok) {
 *   console.log(`Has morph transition: ${result.data}`);
 * }
 */
export async function hasMorphTransition(
  filePath: string,
  slideIndex: number,
): Promise<Result<{ hasMorph: boolean; morphType?: string }>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code ?? "load_failed", zipResult.error?.message ?? "Failed to load presentation");
  }
  const zip = zipResult.data;
  if (!zip) {
    return err("operation_failed", "Failed to load presentation");
  }

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
  }

  const slideEntry = slidePathResult.data;
  if (!slideEntry) {
    return err("slide_not_found", "Slide entry not found");
  }
  const slideXml = requireEntry(zip, slideEntry);

  // Check for morph in transition properties
  const transitionMatch = slideXml.match(/<p:transition\b[\s\S]*?<\/p:transition>/i);
  if (transitionMatch) {
    const transitionXml = transitionMatch[0];
    if (transitionXml.includes('type="morph"') || transitionXml.includes("p:morph")) {
      const morphTypeMatch = transitionXml.match(/morphType="([^"]+)"/i);
      return ok({
        hasMorph: true,
        morphType: morphTypeMatch?.[1],
      });
    }
  }

  // Check timing for morph effects
  const timingMatch = slideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/i);
  if (timingMatch) {
    if (timingMatch[0].includes("<p:morph")) {
      return ok({ hasMorph: true });
    }
  }

  return ok({ hasMorph: false });
}

/**
 * Applies a morph transition to a slide.
 * The morph transition creates a smooth animation effect when moving
 * from this slide to the next slide in the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @param spec - Morph specification (type, duration, ease)
 * @returns Result indicating success
 *
 * @example
 * // Apply object morph transition to slide 1
 * const result = await applyMorphTransition("/path/to/presentation.pptx", 1, {
 *   type: "object",
 *   duration: 1000
 * });
 *
 * // Apply text morph for text transformations
 * const result = await applyMorphTransition("/path/to/presentation.pptx", 2, {
 *   type: "text",
 *   duration: 800,
 *   easein: 20
 * });
 */
export async function applyMorphTransition(
  filePath: string,
  slideIndex: number,
  spec?: MorphSpec,
): Promise<Result<void>> {
  const morphSpec = spec ?? { type: "object", duration: 1000 };

  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);

    // Reset ID counter
    const timingMatch = slideXml.match(/<p:timing>[\s\S]*?<\/p:timing>/);
    if (timingMatch) {
      resetMorphIdCounter(timingMatch[0]);
    } else {
      nextMorphId = 1;
    }

    // Build morph transition XML
    const morphType = morphSpec.type ?? "object";
    const duration = morphSpec.duration ?? 1000;
    const easein = morphSpec.easein ?? 0;
    const easeout = morphSpec.easeout ?? 0;

    // Map morph type to OOXML morphType
    const morphTypeMap: Record<string, string> = {
      object: "obj",
      text: "txt",
      chart: "chart",
      table: "tbl",
      picture: "pict",
    };
    const ooxmlMorphType = morphTypeMap[morphType] ?? "obj";

    // Build transition XML
    let transitionXml = `<p:transition spd="med">`;
    transitionXml += `<p:morph morphType="${ooxmlMorphType}"`;
    if (duration !== 1000) {
      transitionXml += ` dur="${duration}"`;
    }
    if (easein > 0) {
      transitionXml += ` accel="${easein * 100}"`;
    }
    if (easeout > 0) {
      transitionXml += ` decel="${easeout * 100}"`;
    }
    transitionXml += `/></p:transition>`;

    // Remove existing transition
    let updatedSlideXml = slideXml.replace(/<p:transition\b[\s\S]*?<\/p:transition>/gi, "");

    // Insert new transition before </p:sld>
    updatedSlideXml = updatedSlideXml.replace(/<\/p:sld>/i, `${transitionXml}
</p:sld>`);

    // If no </p:sld> found, try </p:sldLayout> or </p:sldMaster>
    if (!updatedSlideXml.includes("</p:sld>")) {
      updatedSlideXml = slideXml.replace(/<\/p:sldMaster>/i, `${transitionXml}
</p:sldMaster>`);
    }

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));
    return ok(void 0);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Removes morph transition from a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result indicating success
 *
 * @example
 * const result = await removeMorphTransition("/path/to/presentation.pptx", 1);
 */
export async function removeMorphTransition(
  filePath: string,
  slideIndex: number,
): Promise<Result<void>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    const slidePathResult = getSlideEntryPath(zip, slideIndex);
    if (!slidePathResult.ok) {
      return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
    }

    const slideEntry = slidePathResult.data;
    if (!slideEntry) {
      return err("slide_not_found", "Slide entry not found");
    }
    const slideXml = requireEntry(zip, slideEntry);

    // Remove morph transition
    const updatedSlideXml = slideXml
      .replace(/<p:transition\b[\s\S]*?<\/p:transition>/gi, "")
      .replace(/<p:morph\b[\s\S]*?\/>/gi, "");

    // Build new zip with updated slide
    const newEntries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of zip.entries()) {
      if (name === slideEntry) {
        newEntries.push({ name, data: Buffer.from(updatedSlideXml, "utf8") });
      } else {
        newEntries.push({ name, data });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));
    return ok(void 0);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets the morph transition settings for a slide.
 *
 * @param filePath - Path to the PPTX file
 * @param slideIndex - 1-based slide index
 * @returns Result with morph settings
 *
 * @example
 * const result = await getMorphTransition("/path/to/presentation.pptx", 1);
 * if (result.ok && result.data.enabled) {
 *   console.log(`Morph type: ${result.data.morphType}`);
 *   console.log(`Duration: ${result.data.duration}`);
 * }
 */
export async function getMorphTransition(
  filePath: string,
  slideIndex: number,
): Promise<Result<{ enabled: boolean; morphType?: string; duration?: number; easein?: number; easeout?: number }>> {
  const zipResult = await loadPresentation(filePath);
  if (!zipResult.ok) {
    return err(zipResult.error?.code ?? "load_failed", zipResult.error?.message ?? "Failed to load presentation");
  }
  const zip = zipResult.data;
  if (!zip) {
    return err("operation_failed", "Failed to load presentation");
  }

  const slidePathResult = getSlideEntryPath(zip, slideIndex);
  if (!slidePathResult.ok) {
    return err(slidePathResult.error?.code ?? "slide_not_found", slidePathResult.error?.message ?? "Failed to get slide path");
  }

  const slideEntry = slidePathResult.data;
  if (!slideEntry) {
    return err("slide_not_found", "Slide entry not found");
  }
  const slideXml = requireEntry(zip, slideEntry);

  // Look for morph transition
  const transitionMatch = slideXml.match(/<p:transition\b[\s\S]*?<\/p:transition>/i);
  if (!transitionMatch) {
    return ok({ enabled: false });
  }

  const transitionXml = transitionMatch[0];
  const morphMatch = transitionXml.match(/<p:morph\b([^>]*)\/?>/i);
  if (!morphMatch) {
    return ok({ enabled: false });
  }

  const morphAttrs = morphMatch[1];

  // Parse morph attributes
  const morphTypeMatch = morphAttrs.match(/morphType="([^"]+)"/i);
  const durationMatch = morphAttrs.match(/dur="(\d+)"/i);
  const accelMatch = morphAttrs.match(/accel="(\d+)"/i);
  const decelMatch = morphAttrs.match(/decel="(\d+)"/i);

  // Map OOXML morph type to our type
  const reverseMorphTypeMap: Record<string, string> = {
    obj: "object",
    txt: "text",
    chart: "chart",
    tbl: "table",
    pict: "picture",
  };

  return ok({
    enabled: true,
    morphType: morphTypeMatch ? reverseMorphTypeMap[morphTypeMatch[1]] ?? morphTypeMatch[1] : "object",
    duration: durationMatch ? parseInt(durationMatch[1], 10) : undefined,
    easein: accelMatch ? parseInt(accelMatch[1], 10) / 100 : undefined,
    easeout: decelMatch ? parseInt(decelMatch[1], 10) / 100 : undefined,
  });
}

// ============================================================================
// Deprecated alias (for backwards compatibility)
// ============================================================================

/**
 * @deprecated Use addAnimation instead
 */
export const setAnimation = addAnimation;
