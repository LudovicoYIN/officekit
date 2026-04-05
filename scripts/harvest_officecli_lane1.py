from __future__ import annotations

import json
import shutil
from dataclasses import asdict, dataclass
from pathlib import Path


@dataclass
class CopiedArtifact:
    category: str
    source: str
    target: str
    bytes: int
    note: str


REPO_ROOT = Path(__file__).resolve().parents[1]
SOURCE_ROOT = REPO_ROOT.parent / "OfficeCLI"
FIXTURE_ROOT = REPO_ROOT / "packages" / "parity-tests" / "fixtures"
TARGET_ROOT = FIXTURE_ROOT / "source"

COPY_ITEMS: list[tuple[str, str, str]] = [
    ("readme", "README.md", "readme/OfficeCLI.README.md"),
    ("ci", ".github/workflows/build.yml", "ci/officecli-build.yml"),
    ("examples", "examples/README.md", "examples/README.md"),
    ("examples", "examples/word/README.md", "examples/word/README.md"),
    ("examples", "examples/word/gen-formulas.sh", "examples/word/gen-formulas.sh"),
    ("examples", "examples/word/gen-complex-tables.sh", "examples/word/gen-complex-tables.sh"),
    ("examples", "examples/word/gen-complex-textbox.sh", "examples/word/gen-complex-textbox.sh"),
    ("examples", "examples/excel/README.md", "examples/excel/README.md"),
    ("examples", "examples/excel/gen-beautiful-charts.sh", "examples/excel/gen-beautiful-charts.sh"),
    ("examples", "examples/excel/gen-charts-demo.sh", "examples/excel/gen-charts-demo.sh"),
    ("examples", "examples/ppt/README.md", "examples/ppt/README.md"),
    ("examples", "examples/ppt/gen-beautiful-pptx.sh", "examples/ppt/gen-beautiful-pptx.sh"),
    ("examples", "examples/ppt/gen-animations-pptx.sh", "examples/ppt/gen-animations-pptx.sh"),
    ("examples", "examples/ppt/gen-video-pptx.py", "examples/ppt/gen-video-pptx.py"),
    ("examples", "examples/ppt/templates/README.md", "examples/ppt/templates/README.md"),
    ("sample-binary", "examples/Alien_Guide.pptx", "examples/binaries/Alien_Guide.pptx"),
    ("sample-binary", "examples/Cat-Secret-Life.pptx", "examples/binaries/Cat-Secret-Life.pptx"),
    ("sample-binary", "examples/budget_review_v2.pptx", "examples/binaries/budget_review_v2.pptx"),
    ("sample-binary", "examples/product_launch_morph.pptx", "examples/binaries/product_launch_morph.pptx"),
]

REFERENCE_ONLY = [
    {
        "category": "large-binary",
        "source": "examples/ppt/outputs/3d-sun.pptx",
        "reason": "33 MB sample retained as source-only reference until the parity harness needs committed heavy media artifacts.",
    },
    {
        "category": "large-binary",
        "source": "examples/ppt/models/sun.glb",
        "reason": "4.1 MB 3D model used by the PowerPoint 3D example; tracked here so the future PPT lane can pull it intentionally.",
    },
    {
        "category": "template-corpus",
        "source": "examples/ppt/templates/styles/",
        "reason": "35-style template library is inventoried here but not duplicated yet to keep this parity bootstrap commit reviewable.",
    },
]

SCENARIOS = [
    {
        "id": "readme-quickstart-ppt",
        "source": "README.md",
        "commands": [
            "officecli create deck.pptx",
            "officecli add deck.pptx / --type slide --prop title=\"Q4 Report\" --prop background=1A1A2E",
            "officecli add deck.pptx '/slide[1]' --type shape --prop text=\"Revenue grew 25%\" --prop x=2cm --prop y=5cm --prop font=Arial --prop size=24 --prop color=FFFFFF",
            "officecli view deck.pptx outline",
            "officecli view deck.pptx html",
            "officecli get deck.pptx '/slide[1]/shape[1]' --json",
        ],
        "verification": ["integration", "differential", "docs"],
    },
    {
        "id": "readme-developer-live-preview",
        "source": "README.md",
        "commands": [
            "officecli create deck.pptx",
            "officecli watch deck.pptx --port 26315",
            "officecli add deck.pptx / --type slide --prop title=\"Hello, World!\"",
        ],
        "verification": ["preview", "e2e", "docs"],
    },
    {
        "id": "ci-smoke-docx",
        "source": ".github/workflows/build.yml",
        "commands": [
            "officecli create test_smoke.docx",
            "officecli add test_smoke.docx /body --type paragraph --prop text=\"Hello from CI\"",
            "officecli get test_smoke.docx /body/p[1]",
        ],
        "verification": ["smoke", "integration", "differential"],
    },
    {
        "id": "examples-word-formulas",
        "source": "examples/word/gen-formulas.sh",
        "commands": ["bash gen-formulas.sh"],
        "verification": ["word", "fixture", "integration"],
    },
    {
        "id": "examples-excel-charts",
        "source": "examples/excel/gen-charts-demo.sh",
        "commands": ["bash gen-charts-demo.sh"],
        "verification": ["excel", "fixture", "integration"],
    },
    {
        "id": "examples-ppt-animations",
        "source": "examples/ppt/gen-animations-pptx.sh",
        "commands": ["bash gen-animations-pptx.sh"],
        "verification": ["ppt", "fixture", "integration"],
    },
]


def copy_items() -> list[CopiedArtifact]:
    copied: list[CopiedArtifact] = []
    for category, source_rel, target_rel in COPY_ITEMS:
        source = SOURCE_ROOT / source_rel
        target = TARGET_ROOT / target_rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        copied.append(
            CopiedArtifact(
                category=category,
                source=source_rel,
                target=str(target.relative_to(REPO_ROOT)),
                bytes=target.stat().st_size,
                note="copied from sibling OfficeCLI source snapshot for parity work",
            )
        )
    return copied


def main() -> None:
    TARGET_ROOT.mkdir(parents=True, exist_ok=True)
    copied = copy_items()
    manifest = {
        "sourceProject": "OfficeCLI",
        "targetProject": "officekit",
        "purpose": "Lane 1 parity bootstrap corpus for the OfficeCLI -> officekit migration",
        "copiedArtifacts": [asdict(item) for item in copied],
        "referenceOnlyArtifacts": REFERENCE_ONLY,
        "scenarios": SCENARIOS,
    }
    manifest_path = FIXTURE_ROOT / "manifest.json"
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {manifest_path.relative_to(REPO_ROOT)} with {len(copied)} copied artifacts and {len(SCENARIOS)} scenario definitions.")


if __name__ == "__main__":
    main()
