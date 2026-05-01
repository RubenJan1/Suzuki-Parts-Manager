# engines/engine_inboeken.py
"""
Suzuki Parts Manager — Inboeken Engine (V17)

Doelen (op basis van jullie workflow + foutpreventie):
- WooCommerce export is LEIDEND (read-only bron)
- Inboeken maakt een OUTPUT-lijst (XLSX/CSV) voor WP All Import
- Categorieën komen uit jullie WooCommerce-structuur (caterogries.json)
  en worden als volledige paden geëxporteerd:
    "Originele onderdelen > 2-takt > GT series > GT750"
- Selecties en detectie: ALTIJD aanvullen/mergen (nooit overschrijven)
- Locatie: spaties blijven behouden (leesbaar zoeken/picken)
- Uitverkocht: stock=0 -> prijs=0 + locatie=""
- Pattern parts: suffix "-p" + categorie "… Pattern parts" als die bestaat

Let op:
- Deze engine kent 2 bronnen:
  1) website_df (Woo export)  [alleen lezen]
  2) own_df (inboek-output)   [wordt aangevuld]

"""

from __future__ import annotations
from datetime import date
from dataclasses import dataclass
from typing import Any, Dict, Optional, List, Tuple
from utils.paths import output_root

import os
import re
import math
import shutil
import datetime as _dt
from datetime import datetime
import json
import uuid

import pandas as pd
import random
import string
import html as _html
from pathlib import Path
import traceback
from utils.paths import output_root


def strip_html(s: str) -> str:
    """Verwijder HTML tags en decode entities (voor WC velden zoals Beschrijving/Locatie)."""
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r"<[^>]+>", " ", s)
    s = _html.unescape(s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

# ============================================================
# CATEGORIEËN: laden uit JSON (WooCommerce structuur)
# ============================================================

_CATEGORIES_JSON_TEXT = r"""[
  {
    "name": "Motoren te koop",
    "children": []
  },
  {
    "name": "Sleutels",
    "children": [
      { "name": "Sleutels met plastic kap", "children": [] },
      { "name": "Sleutels zonder plastic kap", "children": [] }
    ]
  },
  {
    "name": "Originele onderdelen",
    "children": [
      {
        "name": "2-takt",
        "children": [
          {
            "name": "T series",
            "children": [
              {
                "name": "T20",
                "children": [
                  { "name": "T20 Documentation", "children": [] },
                  { "name": "T20 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "T125",
                "children": [
                  { "name": "T125 Documentation", "children": [] },
                  { "name": "T125 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "T200",
                "children": [
                  { "name": "T200 Documentation", "children": [] },
                  { "name": "T200 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "T250",
                "children": [
                  { "name": "T250 Documentation", "children": [] },
                  { "name": "T250 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "T350",
                "children": [
                  { "name": "T350 Documentation", "children": [] },
                  { "name": "T350 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "T500",
                "children": [
                  { "name": "T500 Documentation", "children": [] },
                  { "name": "T500 Pattern parts", "children": [] }
                ]
              }
            ]
          },
          {
            "name": "GT series",
            "children": [
              {
                "name": "GT125",
                "children": [
                  { "name": "GT125 Documentation", "children": [] },
                  { "name": "GT125 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GT185",
                "children": [
                  { "name": "GT185 Documentation", "children": [] },
                  { "name": "GT185 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GT250",
                "children": [
                  { "name": "GT250 Documentation", "children": [] },
                  { "name": "GT250 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "X5 / X7",
                "children": [
                  {
                    "name": "X5",
                    "children": [
                      { "name": "X5 Documentation", "children": [] },
                      { "name": "X5 Pattern parts", "children": [] }
                    ]
                  },
                  {
                    "name": "X7",
                    "children": [
                      { "name": "X7 Documentation", "children": [] },
                      { "name": "X7 Pattern parts", "children": [] }
                    ]
                  }
                ]
              },
              {
                "name": "GT380",
                "children": [
                  { "name": "GT380 Documentation", "children": [] },
                  { "name": "GT380 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GT500",
                "children": [
                  { "name": "GT500 Documentation", "children": [] },
                  { "name": "GT500 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GT550",
                "children": [
                  { "name": "GT550 Documentation", "children": [] },
                  { "name": "GT550 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GT750",
                "children": [
                  { "name": "GT750 Documentation", "children": [] },
                  { "name": "GT750 Pattern parts", "children": [] }
                ]
              }
            ]
          },
          {
            "name": "RE-5",
            "children": [
              { "name": "RE-5 Documentation", "children": [] },
              { "name": "RE-5 Pattern parts", "children": [] }
            ]
          },
          {
            "name": "RV",
            "children": [
              { "name": "RV Documentation", "children": [] },
              { "name": "RV Pattern parts", "children": [] }
            ]
          },
          {
            "name": "RG/RGV series",
            "children": [
              {
                "name": "RG125",
                "children": [{ "name": "RG125 Documentation", "children": [] }]
              },
              {
                "name": "RG250",
                "children": [
                  { "name": "RG250 Documentation", "children": [] },
                  { "name": "RG250 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "RGV250",
                "children": [
                  { "name": "RGV250 Documentation", "children": [] },
                  { "name": "RGV250 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "RG500",
                "children": [
                  { "name": "RG500 Documentation", "children": [] },
                  { "name": "RG500 Pattern parts", "children": [] }
                ]
              }
            ]
          },
          {
            "name": "TC",
            "children": [
              { "name": "TC Documentation", "children": [] },
              { "name": "TC Pattern parts", "children": [] }
            ]
          },
          {
            "name": "TM",
            "children": [
              { "name": "TM Documentation", "children": [] },
              { "name": "TM Pattern parts", "children": [] }
            ]
          },
          {
            "name": "TS",
            "children": [
              { "name": "TS Documentation", "children": [] },
              { "name": "TS Pattern parts", "children": [] },
              { "name": "TS100 / TS125 / TS125X", "children": [] },
              { "name": "TS185", "children": [] },
              {
                "name": "TS250 / TS250X",
                "children": [
                  { "name": "TS250 / TS250X Pattern parts", "children": [] }
                ]
              },
              { "name": "TS400", "children": [] },
              { "name": "TS75 / TS80", "children": [] },
              { "name": "TS90", "children": [] }
            ]
          },
          {
            "name": "RM",
            "children": [
              {
                "name": "RM100 / 125 / 185",
                "children": [
                  { "name": "RM100 / 125 / 185 Documentation", "children": [] },
                  { "name": "RM100 / 125 / 185 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "RM250 / RM370 / RM400 / RM450 / RM465 / RM500",
                "children": [
                  {
                    "name": "RM250 / RM370 / RM400 / RM450 / RM465 / RM500 Documentation",
                    "children": []
                  },
                  {
                    "name": "RM250 / RM370 / RM400 / RM450 / RM465 / RM500 Pattern parts",
                    "children": []
                  }
                ]
              },
              {
                "name": "RM50 / RM60 / RM80",
                "children": [
                  { "name": "RM50 / RM60 / RM80 Documentation", "children": [] },
                  { "name": "RM50 / RM60 / RM80 Pattern parts", "children": [] }
                ]
              }
            ]
          },
          {
            "name": "PE",
            "children": [
              { "name": "PE Documentation", "children": [] },
              { "name": "PE Pattern parts", "children": [] }
            ]
          },
          {
            "name": "50cc",
            "children": [
              {
                "name": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
                "children": [
                  {
                    "name": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P Documentation",
                    "children": []
                  },
                  {
                    "name": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P Pattern parts",
                    "children": []
                  }
                ]
              },
              {
                "name": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
                "children": [
                  {
                    "name": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W Documentation",
                    "children": []
                  }
                ]
              },
              {
                "name": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
                "children": [
                  {
                    "name": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50 Documentation",
                    "children": []
                  }
                ]
              },
              {
                "name": "TS50 / TS50X / TS50ER",
                "children": [
                  { "name": "TS50 / TS50X / TS50ER Documentation", "children": [] },
                  { "name": "TS50 / TS50X / TS50ER Pattern parts", "children": [] }
                ]
              }
            ]
          },
          { "name": "80 CC", "children": [] },
          {
            "name": "100 CC",
            "children": [{ "name": "100 CC Documentation", "children": [] }]
          },
          { "name": "B series", "children": [] },
          {
            "name": "Diverse modellen",
            "children": [
              { "name": "Documentation", "children": [] },
              {
                "name": "TF",
                "children": [{ "name": "TF Documentation", "children": [] }]
              },
              {
                "name": "GP",
                "children": [
                  { "name": "GP Documentation", "children": [] },
                  { "name": "GP Pattern parts", "children": [] }
                ]
              },
              { "name": "K10 / 15P / K125 / S32", "children": [] },
              { "name": "M15", "children": [] },
              {
                "name": "RL",
                "children": [{ "name": "RL Pattern parts", "children": [] }]
              },
              { "name": "TR", "children": [] }
            ]
          }
        ]
      },
      {
        "name": "4-takt",
        "children": [
          {
            "name": "GS series",
            "children": [
              {
                "name": "GS250",
                "children": [
                  { "name": "GS250 Documentation", "children": [] },
                  { "name": "GS250 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GS400 / 425 / 450",
                "children": [
                  { "name": "GS400 / 425 / 450 Documentation", "children": [] },
                  { "name": "GS400 / 425 / 450 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GS500",
                "children": [
                  { "name": "GS500 Documentation", "children": [] },
                  { "name": "GS500 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GS550",
                "children": [
                  { "name": "GS550 Documentation", "children": [] },
                  { "name": "GS550 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GS650",
                "children": [
                  { "name": "GS650 Documentation", "children": [] },
                  { "name": "GS650 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GS750",
                "children": [
                  { "name": "GS750 Documentation", "children": [] },
                  { "name": "GS750 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GS850",
                "children": [
                  { "name": "GS850 Documentation", "children": [] },
                  { "name": "GS850 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GS1000",
                "children": [
                  { "name": "GS1000 Documentation", "children": [] },
                  { "name": "GS1000 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GS1100",
                "children": [
                  { "name": "GS1100 Documentation", "children": [] },
                  { "name": "GS1100 Pattern parts", "children": [] }
                ]
              }
            ]
          },
          {
            "name": "GSX series",
            "children": [
              {
                "name": "GSX250",
                "children": [
                  { "name": "GSX250 Documentation", "children": [] },
                  { "name": "GSX250 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GSX 400",
                "children": [
                  { "name": "GSX 400 Documentation", "children": [] },
                  { "name": "GSX 400 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GSX550",
                "children": [
                  { "name": "GSX550 Documentation", "children": [] },
                  { "name": "GSX550 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GSX750",
                "children": [
                  { "name": "GSX750 Documentation", "children": [] },
                  { "name": "GSX750 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GSX1100",
                "children": [
                  { "name": "GSX1100 Documentation", "children": [] },
                  { "name": "GSX1100 Pattern parts", "children": [] }
                ]
              }
            ]
          },
          {
            "name": "GSX-R series",
            "children": [
              { "name": "GSX-R 1000", "children": [] },
              { "name": "GSX-R 600", "children": [] },
              {
                "name": "GSX-R 750",
                "children": [
                  { "name": "GSX-R 750 Documentation", "children": [] },
                  { "name": "GSX-R 750 Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GSX-R 1100",
                "children": [
                  { "name": "GSX-R 1100 Documentation", "children": [] },
                  { "name": "GSX-R 1100 Pattern parts", "children": [] }
                ]
              }
            ]
          },
          {
            "name": "400 / 425 / 450 cc",
            "children": [
              { "name": "400 / 425 / 450 cc Documentation", "children": [] },
              { "name": "400 / 425 / 450 cc Pattern parts", "children": [] }
            ]
          },
          {
            "name": "CS",
            "children": [{ "name": "CS Documentation", "children": [] }]
          },
          {
            "name": "GN",
            "children": [
              { "name": "GN Documentation", "children": [] },
              { "name": "GN Pattern parts", "children": [] }
            ]
          },
          {
            "name": "GR650",
            "children": [
              { "name": "GR650 Documentation", "children": [] },
              { "name": "GR650 Pattern parts", "children": [] }
            ]
          },
          {
            "name": "GV",
            "children": [
              { "name": "GV Documentation", "children": [] },
              { "name": "GV Pattern parts", "children": [] }
            ]
          },
          {
            "name": "LS",
            "children": [
              { "name": "LS Documentation", "children": [] },
              { "name": "LS Pattern parts", "children": [] }
            ]
          },
          {
            "name": "LT",
            "children": [
              { "name": "LT Documentation", "children": [] },
              { "name": "LT Pattern parts", "children": [] }
            ]
          },
          {
            "name": "DR / SP",
            "children": [
              { "name": "DR / SP Documentation", "children": [] },
              { "name": "DR / SP Pattern parts", "children": [] }
            ]
          },
          {
            "name": "VL / VS / VX / VZ",
            "children": [
              { "name": "VL / VS / VX / VZ Documentation", "children": [] },
              { "name": "VL / VS / VX / VZ Pattern parts", "children": [] }
            ]
          },
          {
            "name": "XN85",
            "children": [
              { "name": "XN85 Documentation", "children": [] },
              { "name": "XN85 Pattern parts", "children": [] }
            ]
          },
          {
            "name": "AN Burgman",
            "children": [{ "name": "AN Burgman Documentation", "children": [] }]
          },
          {
            "name": "Unknown / Onbekend",
            "children": [
              { "name": "Unknown / Onbekend Documentation", "children": [] }
            ]
          },
          {
            "name": "Diverse modellen",
            "children": [
              { "name": "Diverse modellen Documentation", "children": [] },
              { "name": "AH", "children": [] },
              { "name": "ALT", "children": [] },
              {
                "name": "DL",
                "children": [
                  { "name": "DL Documentation", "children": [] },
                  { "name": "DL Pattern parts", "children": [] }
                ]
              },
              { "name": "DR-Z", "children": [] },
              {
                "name": "GS",
                "children": [
                  { "name": "GS Documentation", "children": [] },
                  { "name": "GS Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GSF",
                "children": [
                  { "name": "GSF Documentation", "children": [] },
                  { "name": "GSF Pattern parts", "children": [] }
                ]
              },
              {
                "name": "GSR",
                "children": [{ "name": "GSR Documentation", "children": [] }]
              },
              {
                "name": "GSX",
                "children": [
                  { "name": "GSX Documentation", "children": [] },
                  { "name": "GSX Pattern parts", "children": [] }
                ]
              },
              { "name": "GZ", "children": [] },
              { "name": "LT-Z", "children": [] },
              {
                "name": "RF",
                "children": [
                  { "name": "RF Documentation", "children": [] },
                  { "name": "RF Pattern parts", "children": [] }
                ]
              },
              { "name": "SB", "children": [] },
              {
                "name": "SV",
                "children": [
                  { "name": "SV Documentation", "children": [] },
                  { "name": "SV Pattern parts", "children": [] }
                ]
              },
              {
                "name": "TL",
                "children": [
                  { "name": "TL Documentation", "children": [] },
                  { "name": "TL Pattern parts", "children": [] }
                ]
              },
              { "name": "TU", "children": [] },
              { "name": "TV", "children": [] },
              { "name": "UC", "children": [] },
              { "name": "UG", "children": [] },
              { "name": "UH", "children": [] },
              { "name": "UX", "children": [] },
              {
                "name": "XF",
                "children": [{ "name": "XF Documentation", "children": [] }]
              }
            ]
          }
        ]
      }
    ]
  },
  {
    "name": "Hot Parts",
    "children": []
  },
  {
    "name": "Verschillende merken",
    "children": [
      { "name": "Honda", "children": [] },
      { "name": "Kawasaki", "children": [] },
      { "name": "Yamaha", "children": [] }
    ]
  }
]
"""

def load_categories_from_json_text(text: str) -> List[dict]:
    try:
        data = json.loads(text)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []




# Build index: full_path_str -> meta, and name->paths
def _build_category_indexes(nodes: List[dict]) -> Tuple[Dict[str, dict], Dict[str, List[str]], Dict[str, List[str]]]:
    by_path: Dict[str, dict] = {}
    name_to_paths: Dict[str, List[str]] = {}
    alias_to_paths: Dict[str, List[str]] = {}

    def extract_model_aliases(name: str) -> List[str]:
        # pakt bv "GS400,425,450" -> ["GS400","GS425","GS450"]
        n = (name or "").upper()
        # normaliseer separators
        n = n.replace("/", ",")
        parts = [p.strip() for p in n.split(",") if p.strip()]

        out: List[str] = []
        for p in parts:
            # split ook op spaties
            for tok in re.split(r"\s+", p):
                tok = tok.strip()
                if not tok:
                    continue
                if re.fullmatch(r"[A-Z]{1,4}\d{2,4}", tok):
                    out.append(tok)
        # extra: geval "GS400,425,450" waarbij 425/450 zonder prefix staan
        if parts and re.fullmatch(r"[A-Z]{1,4}\d{2,4}", parts[0]):
            prefix = re.match(r"^([A-Z]{1,4})\d{2,4}$", parts[0]).group(1)
            for p in parts[1:]:
                if p.isdigit():
                    out.append(prefix + p)
        return sorted(set(out))

    def rec(node: dict, prefix: List[str]):
        name = str(node.get("name", "")).strip()
        if not name:
            return
        children = node.get("children", []) or []
        path_list = prefix + [name]
        path_str = " > ".join(path_list)

        by_path[path_str] = {"name": name, "path": path_list, "children": children}
        name_to_paths.setdefault(name.upper(), []).append(path_str)

        for alias in extract_model_aliases(name):
            alias_to_paths.setdefault(alias, []).append(path_str)

        for c in children:
            rec(c, path_list)

    for n in nodes:
        rec(n, [])

    return by_path, name_to_paths, alias_to_paths


def load_categories_tree() -> List[dict]:
    candidates = [
        os.path.join(os.getcwd(), "product_categorieen.json"),
        os.path.join(os.getcwd(), "assets", "product_categorieen.json"),
        os.path.join(os.path.dirname(__file__), "..", "assets", "product_categorieen.json"),
    ]
    for p in candidates:
        p = os.path.normpath(p)
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list):
                    return data
            except Exception:
                pass
    return load_categories_from_json_text(_CATEGORIES_JSON_TEXT)

CATEGORIES_TREE = load_categories_tree()
CATEGORY_BY_PATH, CATEGORY_NAME_TO_PATHS, MODEL_ALIAS_TO_PATHS = _build_category_indexes(CATEGORIES_TREE)

def best_category_path_for_name(name: str) -> Optional[str]:
    """
    Kies het beste categorie-pad voor een modelnaam.

    Belangrijk:
    - We willen NIET automatisch "Documentation" of "Pattern parts" pakken.
    - Dus: eerst proberen een "model node" (niet-doc/pattern) te kiezen.
    """
    key = str(name).strip().upper()

    paths = CATEGORY_NAME_TO_PATHS.get(key, [])
    if not paths:
        paths = MODEL_ALIAS_TO_PATHS.get(key, [])
    if not paths:
        return None

    def is_bad_leaf(p: str) -> bool:
        # Alles wat onder Documentation/Pattern parts hangt, willen we niet auto-selecten
        pl = p.lower()
        return ("documentation" in pl) or ("pattern parts" in pl) or ("pattern part" in pl)

    # 1) Eerst: kies diepste pad dat GEEN doc/pattern is
    good = [p for p in paths if not is_bad_leaf(p)]
    if good:
        return sorted(good, key=lambda p: (-p.count(">"), len(p)))[0]

    # 2) Als er echt alleen doc/pattern bestaat, pak dan het "minst diepe" daarvan
    # (meestal de doc/pattern node zelf, maar liever niet nog dieper)
    return sorted(paths, key=lambda p: (p.count(">"), len(p)))[0]






def find_pattern_child_path(model_path: str) -> Optional[str]:
    """Als een model een child heeft die eindigt op 'Pattern parts', return dat pad."""
    node = CATEGORY_BY_PATH.get(model_path)
    if not node:
        return None
    children = node.get("children", []) or []
    for c in children:
        nm = str(c.get("name","")).strip()
        if nm.lower().endswith("pattern parts"):
            return model_path + " > " + nm
    return None

# ============================================================
# Helpers
# ============================================================

SEARCH_LIMIT_DEFAULT = 500

def clean_text(v: Any) -> str:
    if v is None:
        return ""
    return str(v).replace("\u00a0", " ").strip()

def try_int(v: Any, default: int = 0) -> int:
    try:
        s = clean_text(v)
        if s == "":
            return default
        return int(float(s.replace(",", ".")))
    except Exception:
        return default

def try_float(v: Any, default: float = 0.0) -> float:
    try:
        s = clean_text(v)
        if s == "":
            return default
        return float(s.replace(",", "."))
    except Exception:
        return default

def round_up_to_5cent(x: float) -> float:
    """Round up to the next €0,05 step. NaN/None/inf -> 0.0."""
    try:
        x = float(x)
    except Exception:
        return 0.0
    if not math.isfinite(x):
        return 0.0
    if x <= 0:
        return 0.0
    return math.ceil((x * 20) - 1e-12) / 20.0

def parse_price(v: Any) -> float:
    s = clean_text(v)
    if not s:
        return 0.0
    s = s.replace("€", "").strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    return try_float(s, 0.0)

def normalize_part_number(s: str) -> str:
    s = (s or "").strip().upper()
    if not s:
        return ""

    s = re.sub(r"\s+", "", s)  # spaties eruit

    # strip bekende trailing suffix: -000
    s = re.sub(r"-000$", "", s)

    return s


def _nodash(s: str) -> str:
    """Verwijder alle streepjes — voor vergelijking zonder streepjes."""
    return s.replace("-", "")


def _pn_candidates(pn: str) -> list[str]:
    """
    Maak zoekvarianten:
    - exact zoals ingevoerd
    - zonder streepjes (09380-20007 → 0938020007 en andersom)
    - zonder '-000' (als aanwezig)
    - met '-000' (als afwezig)
    """
    pn = normalize_part_number(pn)
    if not pn:
        return []

    out = {pn, _nodash(pn)}

    # strip -000 suffix
    if re.search(r"-0{3}$", pn):
        base = re.sub(r"-0{3}$", "", pn)
        out.add(base)
        out.add(_nodash(base))

    # add -000 suffix
    if not re.search(r"-\d{3}$", pn):
        out.add(pn + "-000")
        out.add(_nodash(pn) + "000")

    return sorted(out)

def normalize_location(loc: Any) -> str:
    """Locatie normalisatie zonder spaties weg te slopen."""
    s = strip_html(clean_text(loc)).upper()
    if not s:
        return ""
    s = re.sub(r"\s+", " ", s).strip()

    m = re.match(r"^(PLD|PD|GT|GS|GB|D|H|Y|K)\s*(\d{1,3})$", s)
    if not m:
        return s

    code = m.group(1)
    num = int(m.group(2))
    if code in {"D", "PD", "PLD"}:
        return f"{code}{num}"
    return f"{code} {num}"

def first_existing_col(df: pd.DataFrame, *candidates: str) -> Optional[str]:
    if df is None or df.empty:
        return None
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        k = str(cand).strip().lower()
        if k in lower_map:
            return lower_map[k]
    return None

def base_model_from_variant(token: str) -> str:
    t = clean_text(token).upper()
    if not t:
        return ""
    t = t.split()[0]
    if "-" in t:
        t = t.split("-")[0]
    m = re.match(r"^([A-Z]+)(\d+)", t)
    if not m:
        return t
    return f"{m.group(1)}{m.group(2)}"

def detect_pattern(title: str, short_desc: str) -> bool:
    t = f"{title} {short_desc}".lower()
    return any(k in t for k in ["pattern", "patern", "patten", "aftermarket"])

# ============================================================
# Dataclasses
# ============================================================

@dataclass
class SearchHit:
    source: str           # "Website" / "Eigen lijst"
    title: str
    short_description: str
    prijs: float
    stock: int
    locatie: str
    categorieen_raw: str
    category_paths: List[str]
    row: Any = None

@dataclass
class AddUpdateResult:
    actie: str            # "Nieuw" / "Update"
    title: str
    stock: int
    prijs: float
    locatie: str
    short_description: str
    productcategorieen: str

class DebugLogger:
    """
    Simpele debug logger:
    - maakt per 'run' een map: output/inboeken/DEBUG/RUN_<rid>
    - schrijft events naar JSONL (1 regel per event)
    - schrijft ook RUNLOG.xlsx (key/value) en STATS.xlsx
    """
    def __init__(self, base_output_dir: str):
        self.base = Path(base_output_dir)
        self.rid = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.debug_dir = self.base / "DEBUG" / f"RUN_{self.rid}"
        self.debug_dir.mkdir(parents=True, exist_ok=True)
        self.events_path = self.debug_dir / f"EVENTS_{self.rid}.jsonl"
        self.stats = {
            "added": 0,
            "updated": 0,
            "errors": 0,
            "pattern_detected": 0,
            "stock_zero_forced": 0,
            "normalized_location_empty": 0,
        }

    def event(self, kind: str, data: dict):
        row = {
            "ts": datetime.now().isoformat(timespec="seconds"),
            "kind": kind,
            **(data or {})
        }
        with open(self.events_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(row, ensure_ascii=False, default=str) + "\n")

    def bump(self, key: str, n: int = 1):
        if key in self.stats:
            self.stats[key] += n

    def write_runlog_xlsx(self, meta: dict):
        try:
            rows = [[k, str(v)] for k, v in (meta or {}).items()]
            pd.DataFrame(rows, columns=["Key", "Value"]).to_excel(self.debug_dir / f"RUNLOG_{self.rid}.xlsx", index=False)
        except Exception:
            pass

    def write_stats_xlsx(self):
        try:
            rows = [[k, int(v)] for k, v in self.stats.items()]
            pd.DataFrame(rows, columns=["Metric", "Value"]).to_excel(self.debug_dir / f"STATS_{self.rid}.xlsx", index=False)
        except Exception:
            pass

# ============================================================
# Engine
# ============================================================

class InboekenEngine:
    def __init__(self):
        self.website_df: pd.DataFrame = pd.DataFrame()
        self._editing_product_id = None
        # eigen lijst (output)
        self.export_leaf_names: bool = True  # True => export alleen leaf-names (GT750|T350)
        self.df: pd.DataFrame = pd.DataFrame(columns=[
            "ID",  # WooCommerce product ID (leeg = nieuw)
            "Title",
            "Productcategorieën",
            "Stock",
            "Short Description",
            "Locatie",
            "Prijs",
            "Prijs_orig",      # voor korting-knoppen
            "Korting_pct",     # 0/11/19
        ])


        self.reiners_df: Optional[pd.DataFrame] = None
        self._load_reiners_if_present()

        self.output_dir = output_root() / "inboeken"
        self.output_dir.mkdir(parents=True, exist_ok=True)

        self.debug = DebugLogger(self.output_dir)
        self.debug.write_runlog_xlsx({
            "engine": "InboekenEngine",
            "run_id": self.debug.rid,
            "output_dir": str(self.output_dir),
        })
        self.debug.event("ENGINE_START", {"output_dir": str(self.output_dir)})

        self._autosave_dir = os.path.join(self.output_dir, "autosave")
        self._daily_autosave_dir = os.path.join(self.output_dir, "daily_autosave")
        os.makedirs(self._autosave_dir, exist_ok=True)
        os.makedirs(self._daily_autosave_dir, exist_ok=True)

        self._dirty = False
        
        self._restore_last_autosave()




    @staticmethod
  
    def _today_str() -> str:
        return date.today().isoformat()  # YYYY-MM-DD

    
    def _autosave_filename(self) -> str:
        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        uid = uuid.uuid4().hex[:6]  # kort maar uniek
        return f"autosave_{ts}_{uid}.xlsx"
    
    def _leaf_names(self, paths: Optional[list[str]] = None) -> str:
        return self._categories_to_export_string(paths or [])

    def _leaf_name(self, path: str) -> str:
        """
        Geef de leaf-naam van een categoriepad terug.
        Voorbeeld: 'Originele onderdelen > 2-takt > RM > RM250 / ...' -> 'RM250 / ...'
        """
        p = clean_text(path)
        if not p:
            return ""
        if ">" not in p:
            return p.strip()
        return p.split(">")[-1].strip()

    


    def _categories_to_export_string(self, paths: list[str]) -> str:
        """Maak de export string voor kolom Productcategorieën.

        - Full-path mode: 'A > B > GT750|A > B > T350'
        - Leaf mode: 'GT750|T350' (handig als WP All Import alleen op leaf-names matcht).
        """
        paths = [clean_text(p) for p in (paths or []) if clean_text(p)]
        if not paths:
            return ""

        if not self.export_leaf_names:
            seen: set[str] = set()
            out: list[str] = []
            for p in paths:
                p_norm = " > ".join([x.strip() for x in p.split(">") if x.strip()])
                if p_norm and p_norm not in seen:
                    seen.add(p_norm)
                    out.append(p_norm)
            return "|".join(out)

        # leaf mode
        names: list[str] = []
        for p in paths:
            nm = self._leaf_name(p)
            if nm:
                names.append(nm)

        use = names

        seen: set[str] = set()
        out: list[str] = []
        for n in use:
            n2 = n.strip()
            if n2 and n2 not in seen:
                seen.add(n2)
                out.append(n2)
        return "|".join(out)
    

    def _prefer_leaf(path: str) -> str:
      node = CATEGORY_BY_PATH.get(path)
      if not node:
          return path
      children = node.get("children", []) or []
      if not children:
          return path

      # voorkeur: Documentation leaf
      for c in children:
          nm = str(c.get("name","")).strip()
          if nm.lower().endswith("documentation"):
              return path + " > " + nm

      # anders: eerste child als fallback
      nm = str(children[0].get("name","")).strip()
      return path + " > " + nm if nm else path
    
    def apply_discount(self, base_price: float, pct: int) -> float:
      try:
          pct = int(pct)
      except Exception:
          pct = 0
      if pct not in (11, 19):
          return round_up_to_5cent(base_price)

      discounted = base_price * (1.0 - (pct / 100.0))
      return round_up_to_5cent(discounted)

    # -----------------------------
    # Website export
    # -----------------------------

    def set_website_df(self, df: pd.DataFrame) -> None:
        self.website_df = df.copy() if df is not None else pd.DataFrame()

    def _extract_hit_from_website_row(self, row: pd.Series) -> SearchHit:
        title_col = first_existing_col(self.website_df, "Naam", "Name", "Title", "post_title", "product_title", "title", "name", "SKU")
        short_col = first_existing_col(self.website_df, "Korte beschrijving", "Short Description", "short_description")
        price_col = first_existing_col(self.website_df, "Reguliere prijs", "Regular price", "Prijs", "price")
        stock_col = first_existing_col(self.website_df, "Voorraad", "Stock", "stock")
        loc_col = first_existing_col(self.website_df, "Locatie", "Location", "Beschrijving", "Description", "beschrijving")
        cat_col = first_existing_col(self.website_df, "Categorieën", "Productcategorieën", "Productcategorieen", "Categories", "Product categories")

        title = clean_text(row.get(title_col, "")) if title_col else ""
        short_desc = strip_html(clean_text(row.get(short_col, ""))) if short_col else ""
        raw_price = parse_price(row.get(price_col, "")) if price_col else 0.0
        prijs = round_up_to_5cent(raw_price)

        stock = try_int(row.get(stock_col, 0))
        locatie = normalize_location(row.get(loc_col, "")) if loc_col else ""
        cats_raw = clean_text(row.get(cat_col, "")) if cat_col else ""

        # categorieen_raw uit WC kan al paths bevatten; we normaliseren naar paths-lijst
        category_paths = self._normalize_any_category_string_to_paths(cats_raw)

        return SearchHit(
            source="Website",
            title=title,
            short_description=short_desc,
            prijs=prijs,
            stock=stock,
            locatie=locatie,
            categorieen_raw=cats_raw,
            category_paths=category_paths,
            row=row
        )

    def _extract_hit_from_own_row(self, row: pd.Series) -> SearchHit:
        title = clean_text(row.get("Title", ""))
        short_desc = clean_text(row.get("Short Description", ""))
        prijs = parse_price(row.get("Prijs", 0.0))
        stock = try_int(row.get("Stock", 0))
        locatie = normalize_location(row.get("Locatie", ""))
        cats_raw = clean_text(row.get("Productcategorieën", ""))

        category_paths = self._normalize_any_category_string_to_paths(cats_raw)

        return SearchHit(
            source="Eigen lijst",
            title=title,
            short_description=short_desc,
            prijs=prijs,
            stock=stock,
            locatie=locatie,
            categorieen_raw=cats_raw,
            category_paths=category_paths,
            row=row
        )

    def _normalize_any_category_string_to_paths(self, raw: str) -> List[str]:
        """Accepteert:
        - 'A|B|C'
        - 'Originele onderdelen > 2-takt > GT series > GT750|...'
        - 'GT750' (losse naam)
        Return: lijst met volledige paden (zodat UI altijd 1 waarheid heeft).
        """
        s = clean_text(raw)
        if not s:
            return []

        parts = re.split(r"\s*[|,;]\s*", s)
        out: List[str] = []
        for p in parts:
            p = p.strip()
            if not p:
                continue
            if ">" in p:
                # al path
                out.append(" > ".join([x.strip() for x in p.split(">") if x.strip()]))
                continue

            # losse naam -> probeer best path
            best = best_category_path_for_name(p)
            out.append(best if best else p)

        # uniq, keep order
        seen = set()
        uniq = []
        for x in out:
            if x not in seen:
                seen.add(x)
                uniq.append(x)
        return uniq

    # -----------------------------
    # Search
    # -----------------------------

    def exact_title_hits(self, title: str, limit: int = SEARCH_LIMIT_DEFAULT) -> List[SearchHit]:
        t    = normalize_part_number(title)
        t_nd = _nodash(t)
        if not t:
            return []

        hits: List[SearchHit] = []

        if not self.df.empty:
            norm = self.df["Title"].astype(str).str.strip().str.upper()
            mask = norm.eq(t) | norm.str.replace("-", "", regex=False).eq(t_nd)
            for _, r in self.df[mask].head(limit).iterrows():
                hits.append(self._extract_hit_from_own_row(r))

        if not self.website_df.empty and len(hits) < limit:
            col = first_existing_col(self.website_df, "Naam", "Name", "Title", "post_title", "product_title", "title", "name", "SKU")
            if col:
                norm = self.website_df[col].astype(str).str.strip().str.upper()
                mask = norm.eq(t) | norm.str.replace("-", "", regex=False).eq(t_nd)
                for _, r in self.website_df[mask].head(limit - len(hits)).iterrows():
                    hits.append(self._extract_hit_from_website_row(r))

        return hits

    def search(self, query: str, limit: int = SEARCH_LIMIT_DEFAULT) -> List[SearchHit]:
        q    = clean_text(query)
        if not q:
            return []
        q_nd = q.replace("-", "")   # variant zonder streepjes

        hits: List[SearchHit] = []

        # eigen lijst eerst
        if not self.df.empty:
            titles = self.df["Title"].astype(str)
            mask = (
                titles.str.contains(q, case=False, na=False) |
                titles.str.replace("-", "", regex=False).str.contains(q_nd, case=False, na=False) |
                self.df["Short Description"].astype(str).str.contains(q, case=False, na=False)
            )
            for _, r in self.df[mask].head(limit).iterrows():
                hits.append(self._extract_hit_from_own_row(r))

        if not self.website_df.empty and len(hits) < limit:
            title_col = first_existing_col(self.website_df, "Naam", "Name", "Title", "post_title", "product_title", "title", "name", "SKU")
            short_col = first_existing_col(self.website_df, "Korte beschrijving", "Short Description", "short_description")

            mask = pd.Series([False] * len(self.website_df), index=self.website_df.index)
            if title_col:
                titles_w = self.website_df[title_col].astype(str)
                mask |= titles_w.str.contains(q, case=False, na=False)
                mask |= titles_w.str.replace("-", "", regex=False).str.contains(q_nd, case=False, na=False)
            if short_col:
                mask |= self.website_df[short_col].astype(str).str.contains(q, case=False, na=False)

            for _, r in self.website_df[mask].head(limit - len(hits)).iterrows():
                hits.append(self._extract_hit_from_website_row(r))

        return hits

    def not_found_status(self, title: str) -> Dict[str, bool]:
        t = normalize_part_number(title)
        in_own = (not self.df.empty) and self.df["Title"].astype(str).str.upper().eq(t).any()
        in_web = False
        if not self.website_df.empty:
            col = first_existing_col(self.website_df, "Naam", "Name", "Title", "post_title", "product_title", "title", "name", "SKU")
            if col:
                in_web = self.website_df[col].astype(str).str.upper().eq(t).any()
        return {"in_own": bool(in_own), "in_web": bool(in_web)}

    # -----------------------------
    # Load -> UI
    # -----------------------------

    def load_product(self, hit: SearchHit) -> Dict[str, Any]:
        return {
            "Title": hit.title,
            "Stock": str(hit.stock),
            "Prijs": f"{hit.prijs:.2f}".replace(".", ","),
            "Locatie": hit.locatie,
            "Short Description": hit.short_description,
            "SelectedCategoryPaths": list(hit.category_paths),
            "Source": hit.source,
        }

    # -----------------------------
    # Add / Update
    # -----------------------------

    def add_or_update(
        self,
        title: str,
        selected_category_paths: List[str],
        stock: Any,
        short_description: str,
        locatie: str,
        prijs: Any,
        wc_id: Optional[str] = None,
        force_new: bool = False,

    ) -> AddUpdateResult:


        title = clean_text(title)
        if not title:
            self.debug.event("BLOCK_SAVE", {"reason": "Title ontbreekt"})
            raise ValueError("Title is verplicht (onderdeelnummer).")
        if str(stock).strip() == "":
            self.debug.event("BLOCK_SAVE", {"reason": "Voorraad ontbreekt"})
            raise ValueError("Voorraad is verplicht (0 mag, maar niet leeg).")

        stock_i = try_int(stock)
        short_desc = clean_text(short_description)
        if stock_i != 0 and not short_desc:
            self.debug.event("BLOCK_SAVE", {"reason": "Korte beschrijving ontbreekt bij voorraad ≠ 0"})
            raise ValueError("Korte beschrijving is verplicht bij voorraad ≠ 0.")


        # categorieën verplicht (UX: tab zet save-knop uit, maar engine blijft streng)
        cats = [clean_text(c) for c in (selected_category_paths or []) if clean_text(c)]
        if not cats:
            self.debug.event("BLOCK_SAVE", {"reason": "Categorieën ontbreken"})
            raise ValueError("Selecteer minstens 1 categorie/model.")

        locatie_n = normalize_location(locatie)
        prijs_f = round_up_to_5cent(parse_price(prijs))

        if stock_i == 0:
            locatie_n = ""
            prijs_f = 0.0
            self.debug.bump("stock_zero_forced")
            self.debug.event("FORCE_OUT_OF_STOCK", {
                "title": title,
                "stock": stock_i,
                "prijs_set": 0.0,
                "locatie_set": ""
            })
        else:
          if not locatie_n:
              self.debug.event("BLOCK_SAVE", {"reason": "Locatie ontbreekt bij voorraad ≠ 0"})
              raise ValueError("Locatie is verplicht bij voorraad ≠ 0.")
          if prijs_f <= 0:
              self.debug.event("BLOCK_SAVE", {"reason": "Prijs ontbreekt of ≤ 0 bij voorraad ≠ 0"})
              raise ValueError("Prijs is verplicht en moet > 0 bij voorraad ≠ 0.")


        # Pattern detectie alleen nog voor logging/categorie-keuze.
        # Belangrijk: nummer zelf mag GEEN -P suffix meer krijgen.
        is_pattern = detect_pattern(title, short_desc)
        if is_pattern:
            self.debug.bump("pattern_detected")
            self.debug.event("PATTERN_DETECTED", {"title": title})

        # Als pattern -> probeer pattern-categorie child te pakken per model-category
        out_paths: List[str] = []
        for p in cats:
            p_norm = " > ".join([x.strip() for x in p.split(">") if x.strip()])
            if is_pattern:
                child = find_pattern_child_path(p_norm)
                out_paths.append(child if child else p_norm)
            else:
                out_paths.append(p_norm)

        # uniq
        seen = set()
        uniq_paths = []
        for p in out_paths:
            if p not in seen:
                seen.add(p)
                uniq_paths.append(p)

        cats_export = self._categories_to_export_string(uniq_paths)

        actie = "Nieuw"
        wc_id_norm = str(wc_id).strip() if wc_id is not None else ""

        if wc_id_norm:
            # UPDATE op ID
            if not self.df.empty and "ID" in self.df.columns:
                mask = self.df["ID"].astype(str).str.strip().eq(wc_id_norm)
                if mask.any():
                    self.df.loc[mask, ["Title", "Productcategorieën", "Stock", "Short Description", "Locatie", "Prijs"]] = [
                        title, cats_export, stock_i, short_desc, locatie_n, prijs_f
                    ]
                else:
                    # nog niet in eigen lijst => toevoegen maar met ID gevuld (update export)
                    self.df = pd.concat([self.df, pd.DataFrame([{
                        "ID": wc_id_norm,
                        "Title": title,
                        "Productcategorieën": cats_export,
                        "Stock": stock_i,
                        "Short Description": short_desc,
                        "Locatie": locatie_n,
                        "Prijs": prijs_f,
                        "Prijs_orig": "",
                        "Korting_pct": "",
                    }])], ignore_index=True)
            else:
                # geen df of geen ID-kolom (zou niet mogen), toch veilig toevoegen
                self.df = pd.concat([self.df, pd.DataFrame([{
                    "ID": wc_id_norm,
                    "Title": title,
                    "Productcategorieën": cats_export,
                    "Stock": stock_i,
                    "Short Description": short_desc,
                    "Locatie": locatie_n,
                    "Prijs": prijs_f,
                    "Prijs_orig": "",
                    "Korting_pct": "",
                }])], ignore_index=True)

            actie = "Update"

        else:
            # GEEN ID => nooit blind op title updaten, behalve exact 1 match zonder ID en force_new=False
            if (not force_new) and (not self.df.empty) and ("Title" in self.df.columns):
                mask_title = self.df["Title"].astype(str).str.strip().str.upper().eq(title.upper())
                if "ID" in self.df.columns:
                    mask_title = mask_title & self.df["ID"].astype(str).str.strip().eq("")
                matches = int(mask_title.sum()) if hasattr(mask_title, "sum") else 0

                if matches == 1:
                    self.df.loc[mask_title, ["Productcategorieën", "Stock", "Short Description", "Locatie", "Prijs"]] = [
                        cats_export, stock_i, short_desc, locatie_n, prijs_f
                    ]
                    actie = "Update"
                # matches 0 of >1 => nieuwe regel (waterproof)

            if actie == "Nieuw":
                self.df.loc[len(self.df)] = {
                    "ID": "",
                    "Title": title,
                    "Productcategorieën": cats_export,
                    "Stock": stock_i,
                    "Short Description": short_desc,
                    "Locatie": locatie_n,
                    "Prijs": prijs_f,
                    "Prijs_orig": "",
                    "Korting_pct": "",
                }


        self._dirty = True
        self._autosave()

        self.debug.event("SAVE_OK", {
            "actie": actie,
            "title": title,
            "wc_id": wc_id_norm,
            "stock": stock_i,
            "prijs": prijs_f,
            "locatie": locatie_n,
            "categories_count": len(uniq_paths),
            "categories_export": cats_export[:200],  # klein houden
            "force_new": bool(force_new),
        })
        if actie == "Nieuw":
            self.debug.bump("added")
        else:
            self.debug.bump("updated")

        self.debug.write_stats_xlsx()
        
        return AddUpdateResult(
            actie=actie,
            title=title,
            stock=stock_i,
            prijs=prijs_f,
            locatie=locatie_n,
            short_description=short_desc,
            productcategorieen=cats_export,
        )
    def _try_restore_from_autosave(self) -> None:
        """Laad de laatste autosave zodat je bij herstart niets kwijt bent."""
        try:
            p_latest = os.path.join(self._autosave_dir, "latest.xlsx")
            if os.path.exists(p_latest):
                df = pd.read_excel(p_latest)
                if df is not None and not df.empty:
                    self.df = df
        except Exception:
            # bewust stil: autosave mag de app niet laten crashen
            pass


    def _autosave(self) -> None:
        if not self._dirty:
            return
        try:
            # unieke autosave (per actie)
            ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            uid = uuid.uuid4().hex[:6]
            p_unique = os.path.join(
                self._autosave_dir,
                f"inboeken_autosave_{ts}_{uid}.xlsx"
            )

            # daily autosave (1 per dag)
            today = _dt.date.today().strftime("%Y-%m-%d")
            p_daily = os.path.join(
                self._daily_autosave_dir,
                f"inboeken_{today}.xlsx"
            )

            self.df.to_excel(p_unique, index=False)
            self.df.to_excel(p_daily, index=False)

            self._dirty = False
        except Exception:
            pass



    def export_output(self, path: str) -> str:
        path = str(path)
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        self.df.to_excel(path, index=False)
        return path

    def clear_autosave(self) -> None:
        """Wis autosave na output opslaan — herstart begint schoon."""
        import glob as _glob

        # Bewaar de meest recente daily autosave als backup vóór wissen
        try:
            daily_files = sorted(
                _glob.glob(os.path.join(self._daily_autosave_dir, "inboeken_*.xlsx")),
                reverse=True,
            )
            if daily_files:
                backup_dir = self.output_dir / "debug"
                backup_dir.mkdir(parents=True, exist_ok=True)
                ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                shutil.copy2(daily_files[0], str(backup_dir / f"daily_autosave_backup_{ts}.xlsx"))
        except Exception:
            pass

        for folder in [self._autosave_dir, self._daily_autosave_dir]:
            for f in _glob.glob(os.path.join(folder, "*.xlsx")):
                try:
                    os.remove(f)
                except Exception:
                    pass
        self.df = pd.DataFrame(columns=[
            "ID", "Title", "Productcategorieën", "Stock", "Short Description",
            "Locatie", "Prijs", "Prijs_orig", "Korting_pct",
        ])
        self._dirty = False



    # -----------------------------
    # Zedder
    # -----------------------------

    def zedder_fill_title_and_desc(self, text: str, current_title: str = "") -> Tuple[Optional[str], Optional[str]]:
        t = clean_text(text)
        if not t:
            return None, None

        part_number = None
        part_description = None

        for line in t.splitlines():
            line = line.strip()
            if not line:
                continue
            m = re.search(r"Part\s*#\s*=\s*([0-9A-Z-]+)", line, flags=re.I)
            if m:
                part_number = m.group(1).strip()
            m = re.search(r"Part\s*Description\s*=\s*(.+)$", line, flags=re.I)
            if m:
                part_description = m.group(1).strip()

        out_title = None
        if part_number:
            cur = clean_text(current_title)
            if not cur:
                out_title = part_number
            else:
                if part_number.replace("-", "") in cur.replace("-", "") or cur.replace("-", "") in part_number.replace("-", ""):
                    out_title = part_number

        return out_title, part_description
    
    def _fallback_category_for_base_model(self, base: str) -> Optional[str]:
        """Als er geen exact model-leaf is, map op prefix naar een algemene categorie."""
        b = (base or "").upper().strip()
        if not b:
            return None

        # 2-takt groepen
        if b.startswith(("RM", "TS", "TC", "TM", "PE")):
            return "Originele onderdelen > 2-takt > " + b[:2]  # bv RM, TS, TC...
        if b.startswith(("GT", "T")):
            return "Originele onderdelen > 2-takt"

        # 4-takt groepen
        if b.startswith(("DR", "SP")):
            return "Originele onderdelen > 4-takt > DR / SP"
        if b.startswith(("GSX", "GS")):
            return "Originele onderdelen > 4-takt > " + ("GSX series" if b.startswith("GSX") else "GS series")
        if b.startswith(("GN", "GR", "GV", "LS", "LT", "VL", "VS", "VX", "VZ", "XN85")):
            # waar mogelijk specifieke node die wél bestaat, anders 'Diverse modellen'
            if b.startswith("GR"):
                return "Originele onderdelen > 4-takt > GR650"
            if b.startswith("XN85"):
                return "Originele onderdelen > 4-takt > XN85"
            return "Originele onderdelen > 4-takt > Diverse modellen"

        # brommers/50cc etc.
        if re.match(r"^(A|AC|AE|ALT|FA|FS|FZ|JR|OR|RB|TS)\d", b):
            return "Originele onderdelen > 2-takt > 50cc"

        return None
    def _map_model_to_existing_category_paths(self, base_model: str) -> List[str]:
        """
        Zet een base model (bv DR650, RM125, LT230) om naar 1 of meer categoriepaden
        die ook echt in jullie Woo boom bestaan.
        """
        bm = (base_model or "").strip().upper()
        if not bm:
            return []

        out: List[str] = []

        # 1) Exacte leaf match (als jullie boom 'DR650' etc echt als node heeft)
        exact = best_category_path_for_name(bm)
        if exact and exact in CATEGORY_BY_PATH:
            out.append(exact)

        # 2) Fallback op prefix -> groepscategorie
        #    LET OP: we voegen alleen toe als dat pad in CATEGORY_BY_PATH zit.
        def add_if_exists(path: str):
            if path and path in CATEGORY_BY_PATH and path not in out:
                out.append(path)

        # 2-takt: series groepen
        if bm.startswith("GT"):
            add_if_exists("Originele onderdelen > 2-takt > GT series")
        if re.match(r"^T\d", bm):
            add_if_exists("Originele onderdelen > 2-takt > T series")
        if bm.startswith("TS"):
            add_if_exists("Originele onderdelen > 2-takt > TS")
        if bm.startswith("RM"):
            add_if_exists("Originele onderdelen > 2-takt > RM")
        if bm.startswith("PE"):
            add_if_exists("Originele onderdelen > 2-takt > PE")
        if bm.startswith("TM"):
            add_if_exists("Originele onderdelen > 2-takt > TM")
        if bm.startswith("TC"):
            add_if_exists("Originele onderdelen > 2-takt > TC")

        # 50cc / mopeds (best-effort)
        if re.match(r"^(A|AC|AE|ALT|FA|FS|FZ|JR|OR|RB|AS)\d", bm):
            add_if_exists("Originele onderdelen > 2-takt > 50cc")

        # 4-takt: series groepen
        if bm.startswith(("DR", "SP")):
            add_if_exists("Originele onderdelen > 4-takt > DR / SP")

        if bm.startswith("GSX"):
            add_if_exists("Originele onderdelen > 4-takt > GSX series")
        elif bm.startswith("GS"):
            add_if_exists("Originele onderdelen > 4-takt > GS series")

        # “diverse modellen” bucket (voor GN/LT/VS/VL/etc)
        if bm.startswith(("GN", "LT", "VL", "VS", "VX", "VZ", "LS", "GV", "GZ", "SV", "TL", "XN")):
            add_if_exists("Originele onderdelen > 4-takt > Diverse modellen")

        # Specifieke singles (als jullie boom ze zo heeft)
        if bm.startswith("GR"):
            add_if_exists("Originele onderdelen > 4-takt > GR650")
        if bm == "XN85":
            add_if_exists("Originele onderdelen > 4-takt > XN85")

        return out
    
    # --- Zedder model -> boom-naam aliases (jullie JSON) ---
    ZEDDER_MODEL_ALIAS = {
        # jullie hebben deze als 1 categorie in de boom
        "GS400": "GS400,425,450",
        "GS425": "GS400,425,450",
        "GS450": "GS400,425,450",

        # voorbeeld: als jullie boom 1 node heeft voor GSX350+GSX400
        "GSX350": "GSX350/400",
        "GSX400": "GSX350/400",

        # als jullie T20/T200 en GT125 wél bestaan, hoeft dit niet,
        # maar kan geen kwaad (alleen nuttig als Zedder soms rare varianten geeft)
        "T20": "T20",
        "T200": "T200",
        "GT125": "GT125",
    }

    MODEL_TO_CAT: Dict[str, str] = {
        # === T series ===
        "T20": "T20",
        "T125": "T125",
        "T200": "T200",
        "T250": "T250",
        "T350": "T350",
        "T500": "T500",

        # === GT series ===
        "GT125": "GT125",
        "GT185": "GT185",
        "GT250": "GT250",
        "GT380": "GT380",
        "GT500": "GT500",
        "GT550": "GT550",
        "GT750": "GT750",

        # === X5 / X7 ===
        "X5": "X5",
        "X7": "X7",

        # === Overige 2-takt top-level ===
        "RE5": "RE-5", "RE5A": "RE-5", "RE5M": "RE-5",
        "RV": "RV",
        "RV125": "RV", "RV90": "RV",
        "RG125": "RG125",
        "RG250": "RG250",
        "RGV250": "RGV250",
        "RG500": "RG500",
        "TC": "TC", "TC100": "TC", "TC120": "TC", "TC125": "TC", "TC185": "TC", "TC305": "TC", "TC90": "TC",
        "TM": "TM", "TM75": "TM", "TM100": "TM", "TM125": "TM", "TM250": "TM", "TM400": "TM",

        # === TS serie ===
        "TS": "TS",  # hoofd-TS documentatie
        "TS50": "TS50 / TS50X / TS50ER", "TS75": "TS75 / TS80", "TS80": "TS75 / TS80",
        "TS90": "TS90", "TS100": "TS100 / TS125 / TS125X", "TS125": "TS100 / TS125 / TS125X",
        "TS125X": "TS100 / TS125 / TS125X", "TS185": "TS185", "TS250": "TS250 / TS250X",
        "TS250X": "TS250 / TS250X", "TS400": "TS400",

        # === RM serie ===
        "RM50": "RM50 / RM60 / RM80", "RM60": "RM50 / RM60 / RM80", "RM80": "RM50 / RM60 / RM80",
        "RM100": "RM100 / 125 / 185", "RM125": "RM100 / 125 / 185", "RM185": "RM100 / 125 / 185",
        "RM250": "RM250 / RM370 / RM400 / RM450 / RM465 / RM500", "RM370": "RM250 / RM370 / RM400 / RM450 / RM465 / RM500",
        "RM400": "RM250 / RM370 / RM400 / RM450 / RM465 / RM500", "RM450": "RM250 / RM370 / RM400 / RM450 / RM465 / RM500",
        "RM465": "RM250 / RM370 / RM400 / RM450 / RM465 / RM500", "RM500": "RM250 / RM370 / RM400 / RM450 / RM465 / RM500",

        # === PE ===
        "PE": "PE", "PE175": "PE", "PE250": "PE", "PE400": "PE",

        # === 50cc ===
        "A50": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "AC50": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "AD50": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "AE50": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "AG50": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "AH50": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "AJ50": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "AP50": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "AR50": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "AS50": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "A50P": "A50 / AC50 / AD50 / AE50 / AG50 / AH50 / AJ50 / AP50 / AR50 / AS50 / A50P",
        "GA50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "HS50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "RB50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "SF50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "TR50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "NM50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "MT50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "ALT50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "AY50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "CF50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "CS50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "CP50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "UF50": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "UX50W": "GA50 / HS50 / RB50 / SF50 / TR50 / NM50 / MT50 / ALT50 / AY50 / CF50 / CS50 / CP50 / UF50 / UX50W",
        "GT50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "F50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "FA50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "FM50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "FR50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "FS50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "FY50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "FZ50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "JR50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "K50P": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "LT50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "RB50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "RG50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "OR50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "ZR50": "GT50 / F50 / FA50 / FM50 / FR50 / FS50 / FY50 / FZ50 / JR50 / K50P / LT50 / RB50 / RG50 / OR50 / ZR50",
        "TS50": "TS50 / TS50X / TS50ER",
        "TS50X": "TS50 / TS50X / TS50ER",
        "TS50ER": "TS50 / TS50X / TS50ER",

        # === 80 CC ===
        "CS80": "80 CC",

        # === 100 CC ===
        "B100": "100 CC",
        "B120": "100 CC",

        # === B series ===
        "B100": "B series",
        "B120": "B series",

        # === Diverse modellen ===
        "DL": "DL",
        "DL1000": "DL",
        "DL1050": "DL",
        "DL650": "DL",
        "DR": "DR / SP",
        "DR125": "DR / SP",
        "DR250": "DR / SP",
        "DR350": "DR / SP",
        "DR600": "DR / SP",
        "DR650": "DR / SP",
        "DR750": "DR / SP",
        "DR800": "DR / SP",
        "SP125": "DR / SP",
        "SP250": "DR / SP",
        "SP370": "DR / SP",
        "SP400": "DR / SP",
        "SP500": "DR / SP",
        "GSF": "GSF",
        "GSF1200": "GSF",
        "GSF1250": "GSF",
        "GSF600": "GSF",
        "GSF650": "GSF",
        "GSF750": "GSF",
        "RF": "RF",
        "RF400": "RF",
        "RF600": "RF",
        "RF900": "RF",
        "SV": "SV",
        "SV1000": "SV",
        "SV650": "SV",
        "XF650": "XF650 / Freewind",

        # === TF ===
        "TF": "TF",
        "TF125": "TF",

        # === GP ===
        "GP": "GP",
        "GP100": "GP",
        "GP125": "GP",

        # === 4-takt ===
        "AN Burgman": "AN Burgman",
        "AN650": "AN Burgman",
        "400 / 425 / 450 cc": "400 / 425 / 450 cc",
        "GS series": "GS series",
        "GS1000": "GS1000",
        "GS1100": "GS1100",
        "GS250": "GS250",
        "GS400 / 425 / 450": "GS400 / 425 / 450",
        "GS550": "GS550",
        "GS650": "GS650",
        "GS750": "GS750",
        "GS850": "GS850",
        "GSX series": "GSX series",
        "GSX 400": "GSX 400",
        "GSX 600": "GSX 600",
        "GSX 650": "GSX 650",
        "GSX 750": "GSX 750",
        "GSX1100": "GSX1100",
        "GSX1200": "GSX1200",
        "GSX1400": "GSX1400",
        "GSX-R series": "GSX-R series",
        "GSX-R 1000": "GSX-R 1000",
        "GSX-R 1100": "GSX-R 1100",
        "GSX-R 600": "GSX-R 600",
        "GSX-R 750": "GSX-R 750",
        "GN": "GN",
        "GN125": "GN",
        "GN250": "GN",
        "GN400": "GN",
        "GR650": "GR650",
        "GSX8R": "GSX8R / GSX8S / GSX900",
        "GSX8S": "GSX8R / GSX8S / GSX900",
        "GSX900": "GSX8R / GSX8S / GSX900",
        "GTX": "GTX",
        "GTX750": "GTX",
        "GX": "GX",
        "GX750": "GX",
        "VL / VS / VX / VZ": "VL / VS / VX / VZ",
        "VL800": "VL / VS / VX / VZ",
        "VS1400": "VL / VS / VX / VZ",
        "VX800": "VL / VS / VX / VZ",
        "VZ800": "VL / VS / VX / VZ",
        "Verschillende merken": "Verschillende merken",
        "Honda": "Verschillende merken > Honda",
        "Kawasaki": "Verschillende merken > Kawasaki",
        "Yamaha": "Verschillende merken > Yamaha",
        "Derbi": "Verschillende merken > Derbi",
        "Onbekend": "Onbekend"
    }

    def _model_to_category_path(self, model_code: str) -> Optional[str]:
        """
        Zet modelcode (RM250, TS125, GS400, etc.) om naar jullie categorie-pad.
        - gebruikt MODEL_TO_CAT mapping
        - vermijdt voorkeur voor Documentation/Pattern parts als er een 'hoofdcategorie' bestaat
        """
        m = clean_text(model_code).upper()
        if not m:
            return None

        # 1) map model -> categorie-naam uit mapping (jij stuurde MODEL_TO_CAT)
        cat_name = self.MODEL_TO_CAT.get(m, m)

        # 2) eerst exact op gemapte categorie-naam
        p = best_category_path_for_name(cat_name)
        if p:
            return p

        # 3) anders exact op model zelf (voor bomen waar model wél bestaat)
        p = best_category_path_for_name(m)
        if p:
            return p

        # 4) laatste redmiddel: fuzzy zoeken op keys (starts-with / contains)
        #    en dan een “beste” kiezen (liefst geen documentation/pattern)
        candidates: list[str] = []
        for key_upper, paths in CATEGORY_NAME_TO_PATHS.items():
            if key_upper.startswith(m) or (m in key_upper):
                candidates.extend(paths)

        if not candidates:
            return None

        def score(path: str) -> tuple:
            leaf = path.split(">")[-1].strip().lower()
            is_bad = ("documentation" in leaf) or ("pattern" in leaf)
            depth = path.count(">")
            return (is_bad, depth, len(path))

        return sorted(set(candidates), key=score)[0]


    def zedder_detect_model_category_paths(self, text: str) -> List[str]:
        """Detecteer modellen uit Zedder en map ze naar jullie Woo categorie-pad (best-effort)."""
        txt = clean_text(text)
        if not txt:
            self.last_zedder_models = set()
            self.last_zedder_unmapped = set()
            return []

        found_names: set[str] = set()
        for line in txt.splitlines():
            line = line.strip()
            if not line:
                continue
            token = line.split()[0]
            base = base_model_from_variant(token)
            if base and re.match(r"^[A-Z]{1,4}\d{2,4}$", base):
                found_names.add(base)

        # debug info (voor UI-log)
        self.last_zedder_models = set(sorted(found_names))
        self.last_zedder_unmapped = set()

        paths: List[str] = []

        def add_if_exists(p: str):
            p_norm = " > ".join([x.strip() for x in p.split(">") if x.strip()])
            if p_norm in CATEGORY_BY_PATH:
                paths.append(p_norm)

        for nm in sorted(found_names):
            # 1) exact leaf match (als die bestaat)
            alias = self.ZEDDER_MODEL_ALIAS.get(nm, nm)
            best = best_category_path_for_name(alias)

            # fallback: als alias niet bestaat maar "nm" wel (of andersom)
            if not best and alias != nm:
                best = best_category_path_for_name(nm)

            if best:
                add_if_exists(best)
                continue

            # 2) prefix/familie mapping (fallbacks)
            m = nm.upper()

            # 2-takt families
            if m.startswith("GT"):
                add_if_exists("Originele onderdelen > 2-takt > GT series")
            elif re.match(r"^T\d", m):
                add_if_exists("Originele onderdelen > 2-takt > T series")
            elif m.startswith("TS"):
                add_if_exists("Originele onderdelen > 2-takt > TS")
            elif m.startswith("TM"):
                add_if_exists("Originele onderdelen > 2-takt > TM")
            elif m.startswith("RM"):
                add_if_exists("Originele onderdelen > 2-takt > RM")
            elif m.startswith("PE"):
                add_if_exists("Originele onderdelen > 2-takt > PE")
            elif m.startswith("TC"):
                add_if_exists("Originele onderdelen > 2-takt > TC")
            elif re.match(r"^(A|AC|AE|AS|FA|FS|FZ|OR|RE|RV|RS|JR|RB)\d", m):
                # scooters / mopeds / misc 2-takt -> jouw boom heeft meestal 50cc bucket
                add_if_exists("Originele onderdelen > 2-takt > 50cc")

            # 4-takt families
            elif m.startswith("DR") or m.startswith("SP") or m.startswith("DS"):
                add_if_exists("Originele onderdelen > 4-takt > DR / SP")
            elif m.startswith("GR650"):
                add_if_exists("Originele onderdelen > 4-takt > GR650")
            elif m.startswith("GSX"):
                add_if_exists("Originele onderdelen > 4-takt > GSX series")
            elif m.startswith("GS"):
                add_if_exists("Originele onderdelen > 4-takt > GS series")
            elif m.startswith("XN85"):
                add_if_exists("Originele onderdelen > 4-takt > XN85")
            else:
                # laatste vangnet (bestaat in jouw boom)
                add_if_exists("Originele onderdelen > 4-takt > Diverse modellen")

            # als er nog steeds niks toegevoegd is, loggen we 'm als unmapped
            # (dit gebeurt bv. als je JSON die fallback-node niet heeft)
            # check: als nm niets heeft bijgedragen:
            # (we doen dat door te kijken of er een nieuw pad bij kwam)
            # simpeler: als exact geen match en geen van de add_if_exists raakte:
            # -> merk als unmapped
            # (we checken dat met best==None + geen path toegevoegd in deze iter)
            # implementatie:
            # NOTE: we kunnen niet makkelijk per-iter diffen zonder extra variabele:
            # daarom hieronder:
            pass
        # uniq (behoud volgorde)
        seen = set()
        uniq = []
        for p in paths:
            if p not in seen:
                seen.add(p)
                uniq.append(p)

        return uniq


    # -----------------------------
    # Reiners
    # -----------------------------

    def _load_reiners_if_present(self) -> None:
        candidates = [
            os.path.join(os.getcwd(), "reiners.xlsx"),
            os.path.join(os.getcwd(), "reiners.csv"),
            os.path.join(os.getcwd(), "GT Lijst xls.xls"),
            os.path.join(os.getcwd(), "assets", "reiners.xlsx"),
            os.path.join(os.getcwd(), "assets", "reiners.csv"),
            os.path.join(os.getcwd(), "assets", "GT Lijst xls.xls"),
        ]


        p = next((c for c in candidates if os.path.exists(c)), None)
        if not p:
            self.reiners_df = None
            return
        try:
            if p.lower().endswith(".csv"):
                self.reiners_df = pd.read_csv(p, dtype=str, low_memory=False)
            else:
                # .xlsx werkt met openpyxl; .xls kan xlrd nodig hebben
                self.reiners_df = pd.read_excel(p, dtype=str)
            self.reiners_df.columns = [str(c).strip() for c in self.reiners_df.columns]
        except Exception:
            self.reiners_df = None


    def reiners_lookup(self, part_number: str) -> Optional[Dict[str, Any]]:
        if self.reiners_df is None or self.reiners_df.empty:
            return None

        pn = normalize_part_number(part_number)
        if not pn:
            return None

        col_part = first_existing_col(self.reiners_df, "TEILNUMMER", "B", "Part", "PartNumber", "Onderdeelnummer")
        col_desc = first_existing_col(self.reiners_df, "BEZEICHNUNG", "D", "Description", "Beschrijving")
        col_price = first_existing_col(self.reiners_df, "VK ne.", "VK ne", "VK", "Price", "Prijs")

        if not col_part:
            return None

        # kandidaten: met/zonder -000 etc.
        cands = _pn_candidates(part_number)
        if not cands:
            return None
        cands_set = set(cands)

        # kolom normaliseren op zelfde manier
        col_series = (
            self.reiners_df[col_part]
            .astype(str)
            .map(normalize_part_number)
        )

        mask = col_series.isin(cands_set)
        if not mask.any():
            return None

        r = self.reiners_df[mask].iloc[0]

        desc = clean_text(r.get(col_desc, "")) if col_desc else ""

        # prijs uitlezen + normaliseren
        price_raw = clean_text(r.get(col_price, "")) if col_price else ""
        price_raw = price_raw.replace("€", "").replace(",", ".").strip()

        # naar float/decimal via jouw helper
        prijs = round_up_to_5cent(parse_price(price_raw))

        # modellen uit beschrijving
        model_names: set[str] = set()
        for tok in re.findall(r"\b[A-Z]{1,4}\d{2,4}(?:-[A-Z])?\b", desc.upper()):
            base = base_model_from_variant(tok)
            if base and re.match(r"^[A-Z]{1,4}\d{2,4}$", base):
                model_names.add(base)

        # map naar paths
        model_paths: List[str] = []
        for nm in sorted(model_names):
            best = self._model_to_category_path(nm)
            if best:
                model_paths.append(best)

        return {"prijs": prijs, "category_paths": model_paths}

    
    def _restore_last_autosave(self) -> None:
        """Laad automatisch de laatste autosave zodat je niet opnieuw begint."""
        import glob

        # 1) "current" autosave
        p_current = os.path.join(self._autosave_dir, "inboeken_autosave.xlsx")
        if os.path.exists(p_current):
            try:
                self.df = pd.read_excel(p_current, dtype=str).fillna("")
                self._dirty = False
                return
            except Exception:
                pass

        # 2) fallback: meest recente daily autosave
        daily_files = sorted(
            glob.glob(os.path.join(self._daily_autosave_dir, "inboeken_*.xlsx")),
            key=lambda p: os.path.getmtime(p),
            reverse=True,
        )
        if daily_files:
            try:
                self.df = pd.read_excel(daily_files[0], dtype=str).fillna("")
                self._dirty = False
            except Exception:
                pass
            
    def _unique_out_filename(self, base: str = "inboeken_output") -> str:
        ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        suf = "".join(random.choices(string.ascii_uppercase + string.digits, k=4))
        return f"{base}_{ts}_{suf}.xlsx"

            


