"""Load and expose ``budget_config.yaml`` as typed Python objects.

Usage::

    from rsu5.config import cfg

    cfg.cost_centers          # dict[str, CostCenter]
    cfg.articles              # dict[int, Article]
    cfg.column_layout("budget-articles", 27)  # list[str]
    cfg.preferred_doc(27)     # "FY27-budget-articles"
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


_CFG_PATH = Path(__file__).resolve().parent.parent / "budget_config.yaml"


@dataclass
class CostCenter:
    code: str
    name: str
    abbrev: str
    grades: str = ""


@dataclass
class Article:
    number: int
    name: str
    allocated_to_schools: bool


@dataclass
class Config:
    """Parsed budget configuration."""

    raw: dict[str, Any] = field(default_factory=dict, repr=False)

    cost_centers: dict[str, CostCenter] = field(default_factory=dict)
    articles: dict[int, Article] = field(default_factory=dict)
    object_categories: dict[str, str] = field(default_factory=dict)
    function_to_article: dict[str, int] = field(default_factory=dict)

    # Account code regexes (compiled)
    acct_dot_re: re.Pattern | None = field(default=None, repr=False)
    acct_hyphen_re: re.Pattern | None = field(default=None, repr=False)

    # Column layouts: doc_type -> { "FY27": [col_names] }
    _column_layouts: dict[str, dict[str, list[str]]] = field(
        default_factory=dict, repr=False
    )
    _line_item_docs: dict[str, list[str]] = field(
        default_factory=dict, repr=False
    )
    _preferred_docs: dict[str, str] = field(
        default_factory=dict, repr=False
    )

    # Manual data sections (raw dicts)
    enrollment: dict = field(default_factory=dict, repr=False)
    revenue: dict = field(default_factory=dict, repr=False)
    tax_rates: dict = field(default_factory=dict, repr=False)
    eps_data: dict = field(default_factory=dict, repr=False)
    budget_totals: dict = field(default_factory=dict, repr=False)
    assumptions: dict = field(default_factory=dict, repr=False)

    def column_layout(self, doc_type: str, fy: int) -> list[str]:
        """Return the ordered column names for *doc_type* in fiscal year *fy*."""
        key = f"FY{fy}"
        layouts = self._column_layouts.get(doc_type, {})
        if key in layouts:
            return layouts[key]
        raise KeyError(
            f"No column layout for doc_type={doc_type!r} fy={fy}. "
            f"Available: {list(layouts)}"
        )

    def line_item_docs(self, fy: int) -> list[str]:
        """CSV prefixes containing line-item data for *fy*."""
        return self._line_item_docs.get(f"FY{fy}", [])

    def preferred_doc(self, fy: int) -> str:
        """The preferred baseline document prefix for *fy*."""
        return self._preferred_docs[f"FY{fy}"]

    def doc_type_from_prefix(self, prefix: str) -> str:
        """Infer doc_type from a CSV filename prefix like ``FY27-budget-articles``."""
        parts = prefix.split("-", 1)
        if len(parts) < 2:
            return "unknown"
        slug = parts[1]
        mapping = {
            "board-adopted-budget": "adopted",
            "citizens-adopted-budget": "citizens",
            "budget-articles": "proposed",
            "budget-worksheet": "worksheet",
        }
        return mapping.get(slug, slug)

    def abbrev_to_code(self, abbrev: str) -> str | None:
        """Map a cost-center abbreviation (e.g. ``"DCS"``) to its code."""
        for code, cc in self.cost_centers.items():
            if cc.abbrev == abbrev:
                return code
        return None

    def article_for_function(self, func_code: str) -> int | None:
        """Best-effort mapping from function code to article number."""
        if func_code in self.function_to_article:
            return self.function_to_article[func_code]
        prefix = func_code[:2] + "00"
        return self.function_to_article.get(prefix)


def _load(path: Path | None = None) -> Config:
    path = path or _CFG_PATH
    with open(path, encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    c = Config(raw=raw)

    for code, info in raw.get("cost_centers", {}).items():
        c.cost_centers[str(code)] = CostCenter(
            code=str(code),
            name=info["name"],
            abbrev=info["abbrev"],
            grades=info.get("grades", ""),
        )

    for num, info in raw.get("articles", {}).items():
        c.articles[int(num)] = Article(
            number=int(num),
            name=info["name"],
            allocated_to_schools=info["allocated_to_schools"],
        )

    c.object_categories = {
        str(k): v for k, v in raw.get("object_categories", {}).items()
    }

    c.function_to_article = {
        str(k): int(v) for k, v in raw.get("function_to_article", {}).items()
    }

    acct = raw.get("account_code", {})
    if "dot_pattern" in acct:
        c.acct_dot_re = re.compile(acct["dot_pattern"])
    if "hyphen_pattern" in acct:
        c.acct_hyphen_re = re.compile(acct["hyphen_pattern"])

    c._column_layouts = raw.get("column_layouts", {})
    c._line_item_docs = raw.get("line_item_documents", {})
    c._preferred_docs = raw.get("preferred_document", {})

    c.enrollment = raw.get("enrollment", {})
    c.revenue = raw.get("revenue", {})
    c.tax_rates = raw.get("tax_rates", {})
    c.eps_data = raw.get("eps_data", {})
    c.budget_totals = {
        str(k): v for k, v in raw.get("budget_totals", {}).items()
    }
    c.assumptions = raw.get("assumptions", {})

    return c


cfg: Config = _load()
