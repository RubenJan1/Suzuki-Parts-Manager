# ============================================================
# engines/engine_website_277.py
# Website / CMS 277 – Voorraad afboeken
#
# INPUT:
# - Centrale WC-export (AppState.wc_df)
# - Eén of meerdere CMS 277-bestellingen (xlsx)
#
# REGELS:
# - Alleen VOORRAAD wordt aangepast
# - Prijzen blijven onaangeraakt
# - Regel-voor-regel afboeken (geen groupby vooraf)
# - Tekorten worden gelogd
# - Output krijgt ALTIJD een run-id (nooit overschrijven)
# ============================================================
import re
import os
from datetime import datetime
import pandas as pd
import html
import json
import shutil
from pathlib import Path
from utils.paths import output_root


CATEGORIES_JSON_TEXT = r"""[
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
PARTNO_RE = re.compile(r"\b\d{5}-\d{5}\b")

def extract_partnos(text: str) -> set[str]:
    return set(PARTNO_RE.findall(str(text or "")))

def _clean_loc_text(loc: str) -> str:
    """Opschonen van locatie: HTML tags + entities + (escaped) newlines eruit, whitespace normaliseren."""
    if not isinstance(loc, str):
        return ""
    s = html.unescape(loc)
    s = re.sub(r"<[^>]+>", " ", s)
    s = s.replace("\\n", " ").replace("\\r", " ").replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip().upper()
    return s

def is_pattern_by_categories(raw_categories: str) -> bool:
    s = str(raw_categories or "").upper()
    return "PATTERN PARTS" in s


def location_sort_key(loc):
    """
    Natuurlijke sortering voor magazijnlocaties.
    - Excel-achtig: D1, D2, ... D10, D11 ...
    - Samengestelde locaties: 'GS 550/D4' -> D4
    - BGT/BGS/B... worden samengevoegd als 1 groep (B)
    """
    s = _clean_loc_text(loc)
    if not s:
        return (999, "", 999999, "")

    if "/" in s:
        parts = [p.strip() for p in s.split("/") if p.strip()]
        for p in reversed(parts):
            if re.search(r"[A-Z]+\s*\d+", p):
                s = p
                break

    s = re.sub(r"\bP\s+D\b", "PD", s)
    s = re.sub(r"\bPL\s*D\b", "PLD", s)
    s = s.replace(" ", "")

    if s.startswith("BGT") or s.startswith("BGS") or s == "B":
        prefix, num = "B", 0
    elif s.startswith("BR"):
        prefix, num = "BR", 0
    else:
        m = re.search(r"^([A-Z]+)(\d+)$", s)
        if m:
            prefix, num = m.group(1), int(m.group(2))
        else:
            m2 = re.search(r"([A-Z]+)(\d+)", s)
            if m2:
                prefix, num = m2.group(1), int(m2.group(2))
            else:
                prefix, num = s, 999999

    priority_map = {
        "BR": 10,
        "B": 20,
        "D": 30,
        "PD": 35,
        "PLD": 40,
        "GT": 50,
        "GS": 60,
        "GR": 65,
        "GTR": 70,
        "GB": 75,
        "H": 80,
        "Y": 90,
        "RGV": 100,
    }
    pri = priority_map.get(prefix, 500)
    return (pri, prefix, num, s)


def clean_location(loc):
    return _clean_loc_text(loc)

def picklist_loc_for_sort(loc: str) -> str:
    s = _clean_loc_text(loc)
    if not s:
        return ""

    s = s.replace("-", "/")
    s = s.replace(" / ", "/").replace("/ ", "/").replace(" /", "/")
    s_nospace = re.sub(r"\s+", "", s)

    # speciale statuslocaties altijd onderaan
    if "PATTERN" in s or "TRADELIST" in s or "NIETGEVONDEN" in s_nospace or "LOCATIEONTBREEKT" in s_nospace:
        return "ZZZ999999"

    # ✅ BELANGRIJK: pak altijd de LAATSTE 'D<number>' die in de tekst voorkomt
    # werkt voor: 'GT 750 D6', '50 CC / D 8', 'GSX 750 D3', 'D 0', etc.
    d_matches = re.findall(r"D\s*(\d+)", s, flags=re.I)
    if d_matches:
        return f"D{int(d_matches[-1])}"

    # B-groep varianten
    if (
        re.search(r"\bB\s*(GT|GS)\b", s)
        or re.search(r"^B/?(GT|GS)$", s_nospace)
        or s_nospace.startswith("BGT")
        or s_nospace.startswith("BGS")
        or s_nospace == "B"
    ):
        if "GT" in s_nospace:
            return "B/GT"
        if "GS" in s_nospace:
            return "B/GS"
        return "B"

    if s_nospace == "BR":
        return "BR"

    # letters+nummers aan het eind (GT36, PLD103, PL0, GB1, etc.)
    m2 = re.search(r"([A-Z]{1,6}\d+)$", s_nospace)
    if m2:
        return m2.group(1)

    return s




# ------------------------------------------------------------
# PATTERN / NAMA(A)K detectie (277)
# ------------------------------------------------------------
_PATTERN_RE = re.compile(r"\b(PATTERN|NAMA\s*A\s*K|NAMA\s*AK|REPRO|REPLICA)\b", re.I)

def is_pattern_product(title: str, short_desc: str, categories: str) -> bool:
    """True als dit een namaak/PATTERN product is (mag NIET afgeboekt worden)."""
    hay = " | ".join([
        str(title or ""),
        str(short_desc or ""),
        str(categories or ""),
    ])
    hay = html.unescape(hay)
    hay = re.sub(r"<[^>]+>", " ", hay)
    hay = re.sub(r"\s+", " ", hay).strip()
    return bool(_PATTERN_RE.search(hay))


def first_leaf_category(raw_categories: str) -> str:
    """Pak de *eerste* categorie-leaf (na '>') uit WC-export categorie string."""
    if not raw_categories:
        return ""
    # WC-export is meestal: "A > B > C, X > Y"
    first = str(raw_categories).split(",")[0].strip()
    if ">" in first:
        return first.split(">")[-1].strip()
    return first.strip()


def location_sort_key_alpha(loc):
    """
    Alphabetische sortering A..Z, daarna nummers (natuurlijke sort).
    Special:
      - B GT / BGS / B/GT / b gt / B-GT -> 1 B-groep zodat ze bij elkaar staan
      - Bij samengestelde locaties met '/', proberen we eerst B/GT-achtige vormen te normaliseren
    """
    s = _clean_loc_text(loc)
    if not s:
        # lege/None locaties altijd onderaan
        return ("ZZZ", 999999, "ZZZ")

    s = s.upper().strip()

    # normaliseer separators/spaties
    s = re.sub(r"\s+", "", s)
    s = s.replace("-", "/")  # tolerant

    # Special: B/GT, B/GS (en varianten) moeten in B-groep blijven
    # Voorbeelden: 'B/GT', 'B/GS', 'B/GT3' (zeldzaam), 'B/GT ' etc.
    if re.match(r"^B/?(GT|GS)$", s):
        # gebruik 'B' als prefix, en zet subtype in sort-string zodat GT/GS stabiel blijft
        subtype = "GT" if "GT" in s else "GS"
        return ("B", 0, f"B{subtype}")

    # Als samengestelde locatie: kies het "meest locatie-achtige" deel.
    # We nemen het deel dat het beste matcht op: letters+nummer (bv GT10, D3, TLC62, PLD103)
    if "/" in s:
        parts = [p for p in s.split("/") if p]
        # eerst proberen: exact B + (GT|GS) in een van de parts
        for p in parts:
            if re.match(r"^B(GT|GS)$", p):
                subtype = "GT" if "GT" in p else "GS"
                return ("B", 0, f"B{subtype}")
        # anders: pak laatste part met letters+nummer, anders laatste part
        chosen = None
        for p in reversed(parts):
            if re.match(r"^[A-Z]{1,6}\d+$", p):  # GT10, D105, TLC62, PLD103
                chosen = p
                break
        s = chosen if chosen else parts[-1]

    # Nog wat normalisaties voor PL D103 / P Dxx etc
    s = re.sub(r"^PLD(\d+)$", r"PLD\1", s)
    s = re.sub(r"^PD(\d+)$", r"PD\1", s)

    # BGT/BGS samennemen in B groep (ook zonder slash)
    if s.startswith("BGT") or s.startswith("BGS") or s == "B":
        subtype = "GT" if s.startswith("BGT") else ("GS" if s.startswith("BGS") else "")
        return ("B", 0, f"B{subtype}")


        # puur nummer
    if re.match(r"^\d+$", s):
        return ("ZZZ", -int(s), s)  # nummers altijd onderaan, maar groot->klein
    # puur nummers ook aflopend



    # letters+nummer (natuurlijke sort)
    m = re.match(r"^([A-Z]+)(\d+)$", s)
    if m:
        prefix = m.group(1)
        num = int(m.group(2))
        return (prefix, -num, s)  # <-- aflopend


    # letters/overig zonder nummer (komt na prefix+nummer binnen dezelfde prefix)
    # bv 'BR', 'CC', 'GTR', 'OS', 'STD', 'TWB'
    return (s, 999999999, s)



MODEL_REGEX = re.compile(
    r"\b(GT\s*\d+|T\s*\d+|TS\s*\d+|RV\s*\d+|RG\s*\d+|RGV\s*\d+|GSX[\-\s]*R?\s*\d+|GSF\s*\d+|GS\s*\d+|DR\s*\d+|SP\s*\d+|RM\s*\d+|VL\s*\d+|VS\s*\d+|VZ\s*\d+|VX\s*\d+|RF\s*\d+)\b",
    re.I
)

def extract_model_from_categories(cat_text: str) -> str:
    """Pak het laatste model uit de categorie-tekst (GS550, GT750, T125, etc.)."""
    s = str(cat_text or "")
    s = html.unescape(s)
    s = re.sub(r"<[^>]+>", " ", s)
    s = s.replace("\\n", " ").replace("\\r", " ").replace("\n", " ").replace("\r", " ")
    matches = MODEL_REGEX.findall(s)
    if not matches:
        return ""
    m = matches[-1]
    m = re.sub(r"\s+", "", m).upper()
    m = m.replace("GSX-R", "GSXR").replace("GSX R", "GSXR").replace("GSX-", "GSX")
    return m

def model_sort_key(model: str):
    m = (model or "").upper().replace(" ", "")
    if not m:
        return (999, "", 999999)
    fam_map = {
        "GT": 10, "T": 12, "TS": 14, "RV": 16, "RGV": 18,
        "GS": 30, "GSX": 32, "GSXR": 33, "GSF": 34,
        "DR": 50, "SP": 52, "RM": 54, "RF": 56,
        "VL": 70, "VS": 72, "VZ": 74, "VX": 76,
    }
    fam = re.match(r"^[A-Z]+", m).group(0)
    pri = fam_map.get(fam, 500)
    num_m = re.search(r"(\d+)", m)
    num = int(num_m.group(1)) if num_m else 999999
    return (pri, fam, num)

def pick_location(locatie: str, raw_categories: str) -> str:
    """
    Nieuwe regel (277):
    - Alleen voor pure D1..D6: zet het MODEL ervoor: "<MODEL> D#"
      (dus NIET '2-takt', maar bv 'RG 500 D1' / 'GT 750 D4')
    - Alle andere locaties blijven zoals ze zijn
    """
    loc = clean_location(locatie)
    if not loc:
        return ""

    # Als locatie al een model + D# bevat (bv "GS 1000 D6"), NIET overschrijven
    m_loc = re.search(r"\b([A-Z]{1,6})\s*(\d{2,4})\s*D([1-6])$", loc, flags=re.I)
    if m_loc:
        fam = m_loc.group(1).upper()
        num = m_loc.group(2)
        d = m_loc.group(3)
        return f"{fam} {num} D{d}"

    loc_norm = re.sub(r"\s+", "", loc).upper()   # 'D 1' -> 'D1', 'RV D2' -> 'RVD2'
    m = re.fullmatch(r"D([1-6])", loc_norm)      # ✅ alleen exact D1..D6
    if not m:
        return loc


    model = extract_model_from_categories_by_set(raw_categories, MODEL_SET)
    if model:
        # maak 'RG500' -> 'RG 500'
        model_pretty = re.sub(r"^([A-Z]+)(\d+)$", r"\1 \2", model)
        return f"{model_pretty} D{m.group(1)}"

    # fallback als er geen model gevonden wordt
    leaf = first_leaf_category(raw_categories)
    if leaf:
        return f"{leaf} D{m.group(1)}"
    return f"D{m.group(1)}"



def clean_text(val: str) -> str:
    """Opschonen van tekstvelden (korte beschrijving e.d.): HTML tags/entities/newlines eruit."""
    if not isinstance(val, str):
        val = "" if val is None else str(val)
    s = html.unescape(val)
    s = re.sub(r"<[^>]+>", " ", s)
    s = s.replace("\\n", " ").replace("\\r", " ").replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s



def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def run_id():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def to_int(x):
    try:
        return int(float(str(x).replace(",", ".")))
    except:
        return 0
    
def parse_price(val):
    if val is None:
        return 0.0

    s = str(val)

    # verwijder euro, spaties, rare tekens
    s = s.replace("€", "").replace(" ", "").strip()

    # NL → EN decimaal
    s = s.replace(",", ".")

    try:
        return float(s)
    except:
        return 0.0
    
def split_categories(raw):
    if not raw:
        return ""
    parts = [p.strip() for p in str(raw).split(",")]
    leaves = []
    for p in parts:
        if ">" in p:
            leaves.append(p.split(">")[-1].strip())
    return "|".join(sorted(set(leaves)))


def build_model_set_from_categories_json(categories_json_text: str) -> set[str]:
    """
    Haal alle modelnamen uit de categorieboom.
    We nemen items die eruit zien als modelcodes: letters + cijfers, bv RG500, GT750, GSX1100.
    """
    data = json.loads(categories_json_text)
    models = set()

    def walk(node):
        name = str(node.get("name", "")).strip()
        # strip extra suffixen
        name_clean = name.replace("GSX-R", "GSXR").replace("GSX-R ", "GSXR ").replace("GSX-R", "GSXR")

        # vind tokens zoals RG500 / GT750 / GSX1100 in de naam
        for tok in re.findall(r"\b[A-Z]{1,6}\d{2,4}\b", name_clean.upper()):
            models.add(tok)

        for ch in node.get("children", []) or []:
            walk(ch)

    for root in data:
        walk(root)

    return models

MODEL_SET = build_model_set_from_categories_json(CATEGORIES_JSON_TEXT)

def extract_model_from_categories_by_set(cat_text: str, model_set: set[str]) -> str:
    s = clean_text(cat_text).upper()
    s = s.replace("GSX-R", "GSXR").replace("GSX R", "GSXR").replace("GSX-", "GSX")
    
    hits = []
    for m in model_set:
        # ✅ match alleen hele tokens
        if re.search(rf"\b{re.escape(m)}\b", s):
            hits.append(m)

    if not hits:
        return ""

    # kies laatste (meest specifieke in pad), maar bij gelijk: langste eerst
    hits.sort(key=lambda x: (s.rfind(x), len(x)))
    return hits[-1]



class Website277Engine:
    def __init__(self, app_state):
        self.app_state = app_state
        self.cms_paths = []
        self.last_invoice_lines: list = []

    # --------------------------------------------------------
    # INPUT
    # --------------------------------------------------------
    def add_cms_277(self, path: str):
        self.cms_paths.append(path)

    def clear(self):
        self.cms_paths = []

    # --------------------------------------------------------
    # RUN
    # --------------------------------------------------------
    def run(self):
        self.last_invoice_lines = []
        if self.app_state.wc_df is None:
            raise RuntimeError("Geen WC-export geladen")
        if not self.cms_paths:
            raise RuntimeError("Geen CMS 277-bestellingen toegevoegd")

        wc = self.app_state.wc_df.copy()
        wc.columns = [str(c).strip() for c in wc.columns]


        # kolommen tolerant vinden
        col_title = self._find_col(wc, ["Title", "Naam", "post_title"])
        col_id = self._find_col(wc, ["ID", "Id", "post_id", "Post ID", "Product ID"])

        col_stock = self._find_col(wc, ["Stock", "Voorraad"])
        col_locatie = self._find_col(wc, ["Beschrijving"])
        col_short = self._find_col(wc, ["Korte beschrijving", "Short description"])
        col_price = self._find_col(wc, ["Reguliere prijs", "Prijs"])
        col_locatie = self._find_col(wc, ["Beschrijving"])
        col_cat = self._find_col(wc, ["Categorieën"])



        wc[col_stock] = wc[col_stock].apply(to_int)
        wc[col_title] = wc[col_title].astype(str).str.strip()
        wc[col_id] = wc[col_id].astype(str).str.strip()

        updates = []
        picklijst = []
        tekort_log = []
        uitverkocht_log = []  # items die (na levering) op 0 komen of al 0 waren
        changes = []



        # ----------------------------------------------------
        # CMS regels
        # - Meerdere uploads worden eerst samengevoegd op Title+Omschrijving+Prijs
        # ----------------------------------------------------
        cms_rows = []

        for p in self.cms_paths:
          # CMS 277 heeft GEEN headers → positioneel
            df = pd.read_excel(p, header=None, dtype=str)

            # CMS 277 vaste structuur:
            # 0 = Titel / Artikelnummer
            # 1 = Omschrijving
            # 2 = Aantal
            # 3 = Prijs
            # 4 = Factuurnummer

            df = df.rename(
                columns={
                    0: "Title",
                    1: "Omschrijving",
                    2: "Aantal",
                    3: "Prijs",
                    4: "Factuur"
                }
            )

            df["Title"] = df["Title"].fillna("").astype(str).str.strip()
            df["Omschrijving"] = df["Omschrijving"].fillna("").astype(str)
            df["Aantal"] = df["Aantal"].apply(to_int)
            df["Prijs"] = df["Prijs"].apply(parse_price)
            df["Factuur"] = df["Factuur"].fillna("").astype(str)

            df = df[df["Title"] != ""]


            df = df[df["Title"] != ""]

            
            cms_rows.extend(df.to_dict('records'))

        # Samengevoegd: dubbele regels -> aantallen optellen
        total_lines_raw = len(cms_rows)
        total_files = len(self.cms_paths)
        cms_df = pd.DataFrame(cms_rows)
        cms_df["Title"] = cms_df["Title"].fillna("").astype(str).str.strip()
        cms_df["Omschrijving"] = cms_df["Omschrijving"].fillna("").astype(str)
        cms_df["Aantal"] = cms_df["Aantal"].apply(to_int)
        cms_df["Prijs"] = cms_df["Prijs"].apply(parse_price)
        cms_df["Factuur"] = cms_df["Factuur"].fillna("").astype(str)

        def _join_fact(series):
            vals = sorted({str(x).strip() for x in series if str(x).strip()})
            return ", ".join(vals)

        cms_df = (
            cms_df.groupby(["Title", "Omschrijving", "Prijs"], as_index=False)
                  .agg({"Aantal": "sum", "Factuur": _join_fact})
        )
        total_lines_grouped = int(len(cms_df))
        duplicates_merged = int(total_lines_raw - total_lines_grouped)

        # ----------------------------------------------------
        # TRADELIST filter: alleen afboeken als Title op tradelist staat
        # Verwacht: self.app_state.tradelist_path -> pad naar .xlsx met kolom "Artikelnummer"
        # ----------------------------------------------------
        trade_set = set()
        try:
            tl_path = getattr(self.app_state, "tradelist_path", None)
            if tl_path and os.path.exists(tl_path):
                tdf = pd.read_excel(tl_path, dtype=str)
                tdf.columns = [str(c).strip() for c in tdf.columns]
                if "Artikelnummer" in tdf.columns:
                    trade_set = set(tdf["Artikelnummer"].fillna("").astype(str).str.strip())
        except Exception:
            trade_set = set()

        found_lines = 0
        not_found_lines = 0
        for _, r in cms_df.iterrows():
            title = r["Title"]
            qty = int(r["Aantal"])
            # Blokkeer alles dat niet op de tradelist staat
            if trade_set and title not in trade_set:
                tekort_log.append([title, 0, qty, qty, r["Factuur"], "Niet op TRADELIST – niet afboeken"])
                picklijst.append([
                    title,
                    r["Omschrijving"],
                    qty,
                    0,
                    0,
                    0,
                    "TRADELIST (GEBLOKKEERD)",
                    r["Prijs"],
                    r["Factuur"],
                    "Niet op TRADELIST – niet afboeken"
                ])
                not_found_lines += 1
                continue

            hit = wc[wc[col_title] == title].copy()
            matched_via = "TITLE"

            # Als exact match bestaat maar alles is 0 voorraad -> behandel alsof niet gevonden,
            # zodat we superseded/same-as via korte beschrijving kunnen proberen.
            if not hit.empty:
                hit["_stock_tmp"] = hit[col_stock].apply(to_int)
                if hit["_stock_tmp"].max() <= 0:
                    hit = hit.iloc[0:0].copy()  # maak empty om fallback te triggeren
                    matched_via = "TITLE_STOCK0"


            # fallback: zoeken in korte beschrijving naar nummer
            if hit.empty:
                mask = (
                    wc[col_short].fillna("").astype(str).str.contains(re.escape(title), case=False, na=False)
                    | wc[col_cat].fillna("").astype(str).str.contains(re.escape(title), case=False, na=False)
                )

                hit = wc[mask].copy()
                matched_via = "SHORTDESC:TITLE"

                if hit.empty:
                    order_nums = extract_partnos(r.get("Omschrijving", ""))
                    for pn in order_nums:
                        mask2 = wc[col_short].fillna("").astype(str).str.contains(re.escape(pn), case=False, na=False)
                        hit2 = wc[mask2].copy()
                        if not hit2.empty:
                            hit = hit2
                            matched_via = f"SHORTDESC:OMS({pn})"
                            break

            # nog steeds niks gevonden
            if hit.empty:
                tekort_log.append([title, 0, qty, qty, r["Factuur"], "Niet gevonden in WC-export (277)"])
                picklijst.append([title, r["Omschrijving"], qty, 0, 0, 0, "NIET GEVONDEN", r["Prijs"], r["Factuur"], "Niet gevonden in WC-export"])
                not_found_lines += 1
                continue

            # meerdere matches via fallback -> niet automatisch afboeken
            if matched_via != "TITLE" and len(hit) > 1:
                tekort_log.append([title, 0, qty, qty, r["Factuur"], f"Meerdere matches via {matched_via} (controle nodig)"])
                picklijst.append([title, r["Omschrijving"], qty, 0, 0, 0, "MEERDERE MATCHES", r["Prijs"], r["Factuur"], f"Meerdere matches via {matched_via} (controle nodig)"])
                not_found_lines += 1
                continue

            found_lines += 1

            # PATTERN detectie: als er meerdere producten met dezelfde Title zijn,
            # boek dan NOOIT af op een PATTERN/namaak-item.
            hit["_is_pattern"] = hit.apply(
                lambda rr: (
                    is_pattern_product(
                        rr.get(col_title, ""),
                        rr.get(col_short, ""),
                        rr.get(col_cat, "")
                    )
                    or is_pattern_by_categories(rr.get(col_cat, ""))
                ),
                axis=1
            )

            non_pattern = hit[hit["_is_pattern"] == False]
            if not non_pattern.empty:
                hit = non_pattern
            else:
                # alles is PATTERN -> nooit afboeken
                tekort_log.append([
                    title,
                    0,
                    qty,
                    qty,
                    r["Factuur"],
                    "PATTERN product – niet afboeken"
                ])
                picklijst.append([
                    title,
                    r["Omschrijving"],
                    qty,
                    0,
                    0,
                    0,
                    "PATTERN (GEBLOKKEERD)",
                    r["Prijs"],
                    r["Factuur"],
                    "PATTERN product – niet afboeken"
                ])
                not_found_lines += 1
                continue



            # kies daarna de rij met meeste voorraad (en dan laagste locatie)
            hit["_stock"] = hit[col_stock].apply(to_int)
            hit["_loc_sort"] = hit[col_locatie].apply(location_sort_key_alpha)
            hit = hit.sort_values(by=["_stock", "_loc_sort"], ascending=[False, True])
            # pak beste short description uit alle hits (bv langste niet-lege)
            short_best = ""
            try:
                shorts = hit[col_short].fillna("").astype(str).tolist()
                shorts_clean = [clean_text(x) for x in shorts]
                shorts_clean = [x for x in shorts_clean if x.strip()]
                if shorts_clean:
                    short_best = max(shorts_clean, key=len)
            except Exception:
                pass


            idx = hit.index[0]
            stock_oud = int(wc.at[idx, col_stock])


            locatie = clean_location(wc.at[idx, col_locatie])
            if not locatie:
                locatie = "Locatie ontbreekt in WC"

            raw_cats = wc.at[idx, col_cat]
            categories_clean = split_categories(raw_cats)
            pick_loc = pick_location(locatie, raw_cats)

            geleverd = min(stock_oud, qty)
            nieuw_stock = stock_oud - geleverd


            tekort = qty - geleverd

            wc.at[idx, col_stock] = nieuw_stock
            prod_id = wc.at[idx, col_id]
            changes.append([
                prod_id,
                title,
                stock_oud,
                qty,
                geleverd,
                nieuw_stock,
                tekort,
                r["Factuur"]
            ])


            # Website update: als voorraad naar 0 gaat -> prijs & locatie leegmaken (alleen in UPDATE-bestand)
            loc_for_update = clean_location(wc.at[idx, col_locatie])
            price_for_update = wc.at[idx, col_price]
            prod_id = wc.at[idx, col_id]
            if int(nieuw_stock) == 0:
                loc_for_update = ""
                price_for_update = ""

            updates.append([
                prod_id,                      # ← NIEUW (WooCommerce ID)
                title,
                categories_clean,
                nieuw_stock,
                short_best or clean_text(wc.at[idx, col_short]),
                loc_for_update,
                price_for_update
            ])

            opm = ""
            if stock_oud == 0:
                opm = "Voorraad was al 0"
            elif geleverd < qty:
                opm = f"Tekort: {qty - geleverd}"

            picklijst.append([
                title,
                r["Omschrijving"],
                qty,
                geleverd,
                stock_oud,
                nieuw_stock,
                pick_loc,
                r["Prijs"],
                r["Factuur"],
                opm
            ])
            if geleverd > 0:
                self.last_invoice_lines.append({
                    "title": title,
                    "omschrijving": str(r["Omschrijving"]),
                    "besteld": int(qty),
                    "geleverd": int(geleverd),
                    "prijs": float(r["Prijs"]) if r["Prijs"] else 0.0,
                    "factuurnummer": str(r["Factuur"]),
                })
            # ---- Uitverkocht debug (277) ----
            # Let op: hier willen we wél de echte prijs/locatie zien, ook al maken we ze leeg in het UPDATE-bestand.
            if stock_oud == 0 or (stock_oud > 0 and nieuw_stock == 0) or (geleverd < qty):
                uitverkocht_log.append([
                    title,
                    stock_oud,
                    qty,
                    tekort,
                    clean_location(wc.at[idx, col_locatie]) or "Locatie ontbreekt in WC",
                    wc.at[idx, col_price],
                    r["Factuur"],
                    opm or ("Uitverkocht (277)" if nieuw_stock == 0 else "Tekort (277)")
                ])
            if geleverd < qty:
                tekort_log.append([
                    title,
                    stock_oud,
                    qty,
                    tekort,
                    r["Factuur"],
                    "Onvoldoende voorraad (277)"
                ])



        # ----------------------------------------------------
        # Output
        # ----------------------------------------------------
        rid = run_id()
        OUT        = output_root() / "277"
        DIR_UPDATE = OUT / "website_update"
        DIR_PICK   = OUT / "picklijsten"
        DIR_TEKORT = OUT / "tekort"
        DEBUG      = OUT / "DEBUG" / f"RUN_{rid}"

        for d in [DIR_UPDATE, DIR_PICK, DEBUG]:
            d.mkdir(parents=True, exist_ok=True)

        # debug: input kopieën
        try:
            if self.app_state.wc_path and os.path.exists(self.app_state.wc_path):
                shutil.copy2(self.app_state.wc_path, str(DEBUG / f"INPUT_WC_{Path(self.app_state.wc_path).name}"))
            for i, src in enumerate(self.cms_paths, start=1):
                shutil.copy2(src, str(DEBUG / f"INPUT_ORDER_{i}_{Path(src).name}"))
        except Exception:
            pass

        # outputpaden
        p_update = str(DIR_UPDATE / f"WEBSITE_UPDATE_277_{rid}.xlsx")
        p_pick   = str(DIR_PICK   / f"WEBSITE_PICKLIJST_277_{rid}.xlsx")
        p_tekort = str(DIR_TEKORT / f"WEBSITE_TEKORT_277_{rid}.xlsx")
        p_out    = str(DEBUG      / f"WEBSITE_UITVERKOCHT_277_{rid}.xlsx")
        p_stats  = str(DEBUG      / f"RUN_STATS_{rid}.xlsx")

        upd_df = pd.DataFrame(
            updates,
            columns=["ID", "Title", "Productcategorieën", "Stock", "Short Description", "Locatie", "Prijs"]
        )
        # Sorteer de export op locatie (A..Z, natuurlijke nummers)
        try:
            upd_df["_loc_sort"] = upd_df["Locatie"].apply(location_sort_key_alpha)
            upd_df = upd_df.sort_values(by=["_loc_sort", "Title"], ascending=True).drop(columns=["_loc_sort"])
        except Exception:
            pass
        upd_df.to_excel(p_update, index=False)



        pick_df = pd.DataFrame(
            picklijst,
            columns=[
                "Title",
                "Omschrijving",
                "Aantal gevraagd",
                "Aantal geleverd",
                "Voorraad oud",
                "Voorraad nieuw",
                "Locatie",
                "Prijs",
                "Factuurnummer",
                "Opmerking"
            ]
        )

        # kolommen hernoemen (baas)
        pick_df = pick_df.rename(columns={
            "Aantal gevraagd": "Besteld",
            "Aantal geleverd": "Levering",
        })

        # Sortering: eerst model (GS550/GT750/...), dan lade/locatie natuurlijk, dan title
        pick_df["_loc_code"] = pick_df["Locatie"].apply(picklist_loc_for_sort)
        pick_df["_loc_sort"] = pick_df["_loc_code"].apply(location_sort_key_alpha)

        pick_df = pick_df.sort_values(
            by=["_loc_sort", "Title"],
            ascending=True
        ).drop(columns=["_loc_code", "_loc_sort"])



        pick_df.to_excel(p_pick, index=False)


        if uitverkocht_log:
            pd.DataFrame(
                uitverkocht_log,
                columns=[
                    "Title", "Voorraad oud", "Besteld", "Tekort",
                    "Locatie", "Prijs", "Factuurnummer", "Reden"
                ]
            ).to_excel(p_out, index=False)
        else:
            p_out = None

        
        pd.DataFrame(
            changes,
            columns=["ID", "Title", "Stock_oud", "Besteld", "Geleverd", "Stock_nieuw", "Tekort", "Factuur"]
        ).to_excel(str(DEBUG / f"CHANGES_277_{rid}.xlsx"), index=False)


        if tekort_log:
            DIR_TEKORT.mkdir(parents=True, exist_ok=True)
            pd.DataFrame(
                tekort_log,
                columns=[
                    "Title",
                    "Voorraad oud",
                    "Besteld",
                    "Tekort",
                    "Factuurnummer",
                    "Reden"
                ]
            ).to_excel(p_tekort, index=False)
        else:
            p_tekort = None

        # ----------------------------
        # Run stats (controle)
        # ----------------------------
        tekort_lines = len(tekort_log)
        p_stats = None

        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            p_stats = str(DEBUG / f"RUN_STATS_{ts}.xlsx")
            stats = [
                ["CMS bestanden", int(total_files)],
                ["Orderregels (raw)", int(total_lines_raw)],
                ["Regels na samenvoegen", int(total_lines_grouped)],
                ["Dubbele regels samengevoegd", int(duplicates_merged)],
                ["Gevonden in WC-export", int(found_lines)],
                ["Niet gevonden in WC-export", int(not_found_lines)],
                ["Regels met tekort", int(tekort_lines)],
            ]
            pd.DataFrame(stats, columns=["Metric", "Value"]).to_excel(p_stats, index=False)
        except Exception:
            pass

        # Kopieer outputbestanden ook naar DEBUG voor volledig runarchief
        for src_pad in [p_update, p_pick]:
            try:
                shutil.copy2(src_pad, str(DEBUG / Path(src_pad).name))
            except Exception:
                pass
        if p_tekort:
            try:
                shutil.copy2(p_tekort, str(DEBUG / Path(p_tekort).name))
            except Exception:
                pass

        paths = [p_update, p_pick]
        if p_tekort:
            paths.append(p_tekort)
        if p_out:
            paths.append(p_out)
        if p_stats:
            paths.append(p_stats)

        return {
            "batch_id": rid,
            "tab": "277",
            "status": "PENDING_IMPORT",
            "update_path": str(p_update),
            "pick_path": str(p_pick),
            "tekort_path": str(p_tekort) if p_tekort else "",
            "stats_path": str(p_stats) if p_stats else "",
            "debug_changes_path": str(DEBUG / f"CHANGES_277_{rid}.xlsx"),
            "wc_path": str(self.app_state.wc_path or ""),
            "cms_paths": [str(p) for p in self.cms_paths],
            "paths": [str(p) for p in paths],
            "merge_source": False,
        }

    # --------------------------------------------------------
    # HELPERS
    # --------------------------------------------------------
    @staticmethod
    def _find_col(df, candidates):
        for c in df.columns:
            if c in candidates:
                return c
        raise RuntimeError(f"Kolom niet gevonden: {candidates}")