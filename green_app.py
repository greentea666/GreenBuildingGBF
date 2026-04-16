# -*- coding: utf-8 -*-
"""
GreenBuildingGBF — Open Source Edition
DWG → GBF Generator for Taiwan Green Building Assessment System v1.5.1

Open source tool: users must provide their own office credentials.
No sensitive data is bundled.
"""

import io
import json
import os
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

if sys.stdout and hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False

try:
    import windnd
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ─────────────────────────────────────────────────────────
# Color palette — Retro OS & Terminal Aesthetics
# Light neutral base + amber/orange accent + thin borders
# ─────────────────────────────────────────────────────────
C = {
    "bg":        "#F5F3EE",    # warm off-white paper
    "bg2":       "#EDEAE4",    # slightly darker paper
    "surface0":  "#E5E2DB",    # card/input background
    "surface1":  "#D5D2CB",    # border light
    "surface2":  "#C5C2BB",    # border medium
    "overlay":   "#8A8780",    # muted caption text
    "text":      "#1A1A1A",    # primary black text
    "subtext":   "#4A4A46",    # secondary dark gray
    "accent":    "#F0A500",    # amber/orange accent (primary)
    "accent2":   "#FFB800",    # lighter amber hover
    "green":     "#2D8A4E",    # success green (muted)
    "teal":      "#1A7A5C",    # teal accent
    "blue":      "#2563EB",    # link blue
    "lavender":  "#6366F1",    # heading accent
    "peach":     "#E8740C",    # warning orange
    "red":       "#DC2626",    # error red
    "yellow":    "#F0A500",    # same as accent
    "pink":      "#DB2777",    # pink accent
    "border":    "#1A1A1A",    # thin black border
    "border_lt": "#B5B2AB",    # light border for sections
}

FONT_TITLE = ("Microsoft JhengHei UI", 20, "bold")
FONT_HEADING = ("Microsoft JhengHei UI", 12, "bold")
FONT_BODY = ("Microsoft JhengHei UI", 10)
FONT_SMALL = ("Microsoft JhengHei UI", 9)
FONT_MONO = ("Consolas", 10)
FONT_MONO_SM = ("Consolas", 9)
FONT_BTN = ("Microsoft JhengHei UI", 11, "bold")
FONT_CAPTION = ("Microsoft JhengHei UI", 8)
FONT_TAG = ("Consolas", 8)           # decorative tags: INPUT, TERMINAL OUTPUT

# ─────────────────────────────────────────────────────────
# Paths & config
# ─────────────────────────────────────────────────────────

def get_app_dir():
    """Get directory where EXE (or .py) lives."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

APP_DIR = get_app_dir()
CONFIG_PATH = os.path.join(APP_DIR, "config.yaml")

DEFAULT_CONFIG = {
    "office": {
        "name": "",
        "tel": "",
        "address": "",
        "visa_holder": "",
        "certificate_no": "",
        "authorization": "",
        "verification": "",
        "term_start": "",
        "term_end": "",
    },
    "window_defaults": {
        "glass": "單層透明玻璃",
        "glass_code": "P5",
        "glass_name": "單層玻璃",
        "color": "平板玻璃",
        "thickness": "3",
        "frame": "鋁門窗窗框",
        "open_ratio": 0.5,
    },
    "site_defaults": {
        "altitude": "< 200m",
        "soil_classification": "回填土",
        "soil_permeability": 0.00001,
        "carbon_benchmark": "前二類以外之建築基地",
    },
    "project_defaults": {
        "apply_type": "建造執照申請（含變更設計）",
        "is_public": 1,
        "land_use_2": "",
        "land_use_3": "",
    },
}


def load_config():
    if not os.path.exists(CONFIG_PATH):
        return dict(DEFAULT_CONFIG)
    try:
        if HAS_YAML:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            # Merge with defaults
            merged = dict(DEFAULT_CONFIG)
            for section in merged:
                if section in data and isinstance(data[section], dict):
                    merged[section] = {**merged[section], **data[section]}
            return merged
        else:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        return dict(DEFAULT_CONFIG)


def save_config(cfg):
    try:
        if HAS_YAML:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                yaml.dump(cfg, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
        else:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showerror("錯誤", f"無法儲存設定檔:\n{e}")


# ─────────────────────────────────────────────────────────
# AutoCAD COM + AutoLISP dump
# ─────────────────────────────────────────────────────────

LISP_DUMP_CODE = r'''(progn (setq f (open "C:/temp/cad_dump.txt" "w")) (vlax-for ent (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object))) (setq oname (vla-get-objectname ent)) (cond ((= oname "AcDbBlockReference") (setq bname (vla-get-name ent)) (setq lay (vla-get-layer ent)) (setq attrs "") (if (= (vla-get-hasattributes ent) :vlax-true) (progn (setq attlist (vlax-invoke ent 'getattributes)) (foreach att attlist (setq attrs (strcat attrs (vla-get-tagstring att) "=" (vla-get-textstring att) "|"))))) (write-line (strcat "BLK~" lay "~" bname "~" attrs) f)) ((or (= oname "AcDbText") (= oname "AcDbMText")) (setq txt (vla-get-textstring ent)) (setq lay (vla-get-layer ent)) (setq ht (rtos (vla-get-height ent) 2 1)) (setq px (rtos (car (vlax-get ent 'insertionpoint)) 2 1)) (setq py (rtos (cadr (vlax-get ent 'insertionpoint)) 2 1)) (write-line (strcat "TXT~" lay "~" ht "~" px "~" py "~" txt) f)))) (close f) (princ "DONE"))'''


def connect_and_dump(dwg_path, log_fn):
    import win32com.client
    dump_path = "C:/temp/cad_dump.txt"
    os.makedirs("C:/temp", exist_ok=True)
    if os.path.exists(dump_path):
        os.remove(dump_path)

    log_fn("[*] 連接 AutoCAD 中...")
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True

    doc = acad.ActiveDocument
    if dwg_path:
        dwg_name = os.path.basename(dwg_path)
        need_open = True
        try:
            if doc and doc.Name == dwg_name:
                need_open = False
        except Exception:
            pass
        if need_open:
            log_fn(f"[*] 開啟圖檔: {dwg_name}")
            acad.Documents.Open(dwg_path)
            time.sleep(2)  # wait for AutoCAD to finish opening
            doc = acad.ActiveDocument

    log_fn(f"[*] 圖面: {doc.Name}")
    log_fn("[*] 掃描中 (約 30-60 秒)...")
    doc.SendCommand(LISP_DUMP_CODE + "\n")

    for i in range(120):
        time.sleep(1)
        if os.path.exists(dump_path):
            s1 = os.path.getsize(dump_path)
            time.sleep(1)
            s2 = os.path.getsize(dump_path)
            if s2 == s1 and s2 > 100:
                break
        if i % 10 == 9:
            log_fn(f"    等待中... ({i+1}秒)")

    if not os.path.exists(dump_path):
        raise RuntimeError("超過 2 分鐘仍未產生資料檔，請確認 AutoCAD 已開啟圖面")
    return dump_path


# ─────────────────────────────────────────────────────────
# Parse dump file
# ─────────────────────────────────────────────────────────

def parse_dump(dump_path):
    with open(dump_path, "rb") as f:
        raw = f.read()
    content = raw.decode("cp950", errors="replace")
    blocks, texts = [], []
    for line in content.splitlines():
        parts = line.split("~")
        if parts[0] == "BLK" and len(parts) >= 4:
            attrs = {}
            if parts[3]:
                for pair in parts[3].split("|"):
                    if "=" in pair:
                        k, v = pair.split("=", 1)
                        attrs[k] = v
            blocks.append({"layer": parts[1], "name": parts[2], "attrs": attrs})
        elif parts[0] == "TXT" and len(parts) >= 6:
            texts.append({
                "layer": parts[1],
                "height": float(parts[2]) if parts[2] else 0,
                "x": float(parts[3]) if parts[3] else 0,
                "y": float(parts[4]) if parts[4] else 0,
                "text": "~".join(parts[5:]),
            })
    return blocks, texts


# ─────────────────────────────────────────────────────────
# Data extraction
# ─────────────────────────────────────────────────────────

def extract_project_info(blocks, texts):
    info = {}
    for blk in blocks:
        if "圖框" in blk["name"]:
            for tag, val in blk["attrs"].items():
                if "圖名" in tag and val:
                    info.setdefault("drawings", []).append(val)
    for t in texts:
        txt = t["text"]
        if re.search(r"[台臺][南北東西]?[市縣].*段\d+號", txt):
            info["address"] = txt.strip()
        if re.match(r"[A-Z]-\d", txt):
            info["use_type"] = txt.strip()
        m = re.search(r"基地面積[（(]?Ao[)）]?[：:]\s*(\d+\.?\d*)\s*[㎡m]", txt)
        if m:
            info["base_area"] = float(m.group(1))
        m = re.search(r"建築面積[：:]\s*(\d+\.?\d*)\s*[㎡m]", txt)
        if m:
            info["building_area"] = float(m.group(1))
        m = re.search(r"謄本面積\s*(\d+\.?\d*)\s*m", txt)
        if m:
            info.setdefault("base_area", float(m.group(1)))
        m = re.search(r"樓地板面積\s*(\d+\.?\d*)\s*平方", txt)
        if m:
            info["floor_area"] = float(m.group(1))
        m = re.search(r"實際建蔽率[：:].*=\s*(\d+\.?\d*)%", txt)
        if m:
            info["coverage_rate"] = float(m.group(1))
        m = re.search(r"綠化面積\s*(\d+\.?\d*)\s*m2", txt, re.IGNORECASE)
        if m:
            val = float(m.group(1))
            if val > info.get("green_area", 0):
                info["green_area"] = val
        m = re.search(r"法定空地面積\s*=?\s*(\d+\.?\d*)\s*m", txt)
        if m:
            info["open_space"] = float(m.group(1))
        m = re.search(r"滲透側溝總長度[：:]\s*([\d.+]+)\s*=\s*([\d.]+)\s*m", txt)
        if m:
            info["infiltration_length"] = float(m.group(2))
        m = re.search(r"滲透陰井[個數：:]*\s*(\d+)\s*個", txt)
        if m:
            info["infiltration_wells"] = int(m.group(1))
    return info


def extract_windows(blocks, texts):
    wno_re = re.compile(r"^(W\d+|DW\d+|D\d+|S\d+)$")
    counts = {}
    for t in texts:
        txt = t["text"].strip()
        if wno_re.match(txt):
            counts[txt] = counts.get(txt, 0) + 1
    for blk in blocks:
        if "窗編號" in blk["name"]:
            wno = blk["attrs"].get("窗編號", "")
            if wno:
                counts[wno] = counts.get(wno, 0) + 1
    return counts


def extract_plants(texts):
    plants = []
    target = [t for t in texts if t["layer"] == "2-戶外地坪"]
    if not target:
        return plants
    target.sort(key=lambda t: -t["y"])
    rows = []
    cur = [target[0]]
    for t in target[1:]:
        if abs(t["y"] - cur[0]["y"]) < 50:
            cur.append(t)
        else:
            rows.append(cur)
            cur = [t]
    rows.append(cur)

    CAT = {
        "草坪": "草花花圃、自然野草地、水生植物、草坪",
        "草皮": "草花花圃、自然野草地、水生植物、草坪",
        "灌木": "灌木(每㎡栽植二株以上)",
        "喬木": "闊葉大喬木",
        "大喬木": "闊葉大喬木",
        "小喬木": "闊葉小喬木、針葉喬木、疏葉喬木",
        "棕櫚": "棕櫚類",
    }
    SKIP = {"A", "B", "C", "D", "A1", "A2", "A3", "高壓混凝土磚"}
    cur_cat = None
    for row in rows:
        txts = [t["text"].strip() for t in row]
        for txt in txts:
            for kw, cat in CAT.items():
                if txt == kw:
                    cur_cat = cat
        if cur_cat is None:
            continue
        area = 0
        for txt in txts:
            m = re.search(r"([\d.]+)\s*[㎡m²]", txt, re.IGNORECASE)
            if m:
                v = float(m.group(1))
                if v > 0.5:
                    area = v
        name = cur_cat
        for txt in txts:
            if txt in SKIP or txt in CAT:
                continue
            if re.search(r"[\d㎡m²]", txt):
                continue
            if 1 < len(txt) < 30:
                name = txt
                break
        if area > 0:
            plants.append({"name": name, "catalog": cur_cat, "area": area})

    unique, seen = [], set()
    for p in plants:
        key = f"{p['catalog']}_{p['area']}"
        if key not in seen:
            unique.append(p)
            seen.add(key)
    return unique


# ─────────────────────────────────────────────────────────
# GBF builder
# ─────────────────────────────────────────────────────────

DATA_PLANT_USER = {
    "闊葉大喬木": {"Gi": 1.5, "Depth": 1, "DepthOther": 1, "Area": 4, "Area2": 1.5,
                    "IsCanNamed": False, "HasNumbers": True, "IsOtherType": False, "HasPlants": True, "Plants": []},
    "闊葉小喬木、針葉喬木、疏葉喬木": {"Gi": 1.0, "Depth": 0.7, "DepthOther": 1, "Area": 1.5, "Area2": 0,
                                          "IsCanNamed": False, "HasNumbers": True, "IsOtherType": False, "HasPlants": True, "Plants": []},
    "棕櫚類": {"Gi": 0.66, "Depth": 0.7, "DepthOther": 1, "Area": 1.5, "Area2": 0,
               "IsCanNamed": False, "HasNumbers": True, "IsOtherType": False, "HasPlants": True, "Plants": []},
    "灌木(每㎡栽植二株以上)": {"Gi": 0.5, "Depth": 0.4, "DepthOther": 0.5, "Area": 0, "Area2": 0,
                                "IsCanNamed": False, "HasNumbers": False, "IsOtherType": False, "HasPlants": True, "Plants": []},
    "多年生蔓藤": {"Gi": 0.4, "Depth": 0.4, "DepthOther": 0.5, "Area": 0, "Area2": 0,
                    "IsCanNamed": False, "HasNumbers": False, "IsOtherType": False, "HasPlants": True, "Plants": []},
    "草花花圃、自然野草地、水生植物、草坪": {"Gi": 0.3, "Depth": 0.1, "DepthOther": 0.3, "Area": 0, "Area2": 0,
                                              "IsCanNamed": False, "HasNumbers": False, "IsOtherType": False, "HasPlants": True, "Plants": []},
    "薄層綠化、壁掛式綠化": {"Gi": 0.3, "Depth": 0.1, "DepthOther": 0.3, "Area": 0, "Area2": 0,
                              "IsCanNamed": False, "HasNumbers": False, "IsOtherType": False, "HasPlants": True, "Plants": []},
    "其他": {"Gi": 0, "Depth": 0, "DepthOther": 0, "Area": 0, "Area2": 0,
             "IsCanNamed": False, "HasNumbers": False, "IsOtherType": True, "HasPlants": True, "Plants": []},
}

USE_TYPE_MAP = {
    "F-3幼兒園": "F-3兒童福利",
}


def make_file_person(cfg=None):
    if cfg is None:
        cfg = {}
    return {
        "Authorization": cfg.get("authorization", ""),
        "Verification": cfg.get("verification", ""),
        "VisaHolder": cfg.get("visa_holder", ""),
        "CertificateNo": cfg.get("certificate_no", ""),
        "OfficeName": cfg.get("name", ""),
        "OfficeTel": cfg.get("tel", ""),
        "OfficeAddress": cfg.get("address", ""),
        "TermStart": cfg.get("term_start", ""),
        "TermEnd": cfg.get("term_end", ""),
    }


def make_energy_saving():
    return {
        "EnergyConsumingPartition1": False, "EnergyConsumingPartition2": False,
        "EnergyConsumingPartition3": False, "EnergyConsumingPartition4": False,
        "EnergyConsumingPartition5": False, "EnergyConsumingPartition6": False,
        "PositionFloorGridData": [], "WindowGridData": [],
        "A1RoofOpaqueGridData": [], "RoofTransparentRfrType": None,
        "B1WallGridData": [], "C1FixedLargeSpaceGridData": [],
        "C1ECPartitionGridData": [], "D3HouseholdWallGridData": [],
        "F1FixedLargeSpaceGridData": [],
    }


def build_gbf(project_info, window_counts, plants, cfg):
    office = cfg.get("office", {})
    person = make_file_person(office)
    win_cfg = cfg.get("window_defaults", {})
    site = cfg.get("site_defaults", {})
    proj = cfg.get("project_defaults", {})

    raw_type = project_info.get("use_type", "F-3兒童福利")
    use_type = USE_TYPE_MAP.get(raw_type, raw_type)
    addr = project_info.get("address", "")

    # Detect city from address
    city = ""
    for c in ["臺北市", "新北市", "桃園市", "臺中市", "臺南市", "高雄市",
              "基隆市", "新竹市", "嘉義市", "新竹縣", "苗栗縣", "彰化縣",
              "南投縣", "雲林縣", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣",
              "臺東縣", "澎湖縣", "金門縣", "連江縣"]:
        if c in addr:
            city = c
            break
    if not city:
        for alias, formal in [("台南", "臺南市"), ("台北", "臺北市"), ("台中", "臺中市"),
                               ("高雄", "高雄市"), ("新北", "新北市"), ("桃園", "桃園市")]:
            if alias in addr:
                city = formal
                break

    gbf = {
        "DataWindowBaseUser": {},
        "AdjacentBuildingData": [],
        "FileCreater": dict(person), "FileWriter": dict(person), "FileOwner": dict(person),
        "ForceShowGreen": False, "ForceShowWater": False, "ForceShowEnergySaving": False,
        "ForceShowGreenBuildMaterial": False, "ForceShowRainwaterStorage": False,
        "BulidingNumber": "", "BulidingName": addr,
        "LandNumberCity": city,
        "LandNumber": addr,
        "ApplyType": proj.get("apply_type", "建造執照申請（含變更設計）"),
        "ConstructionLicenseNo": "",
        "LandUse1": "土地使用分區",
        "LandUse2": proj.get("land_use_2", ""),
        "LandUse3": proj.get("land_use_3", ""),
        "Maker": "", "Designer": "",
        "VisaHolder": person.get("VisaHolder", ""),
        "CertificateNo": person.get("CertificateNo", ""),
        "OfficeName": person.get("OfficeName", ""),
        "OfficeTel": person.get("OfficeTel", ""),
        "OfficeAddress": person.get("OfficeAddress", ""),
        "Altitude": site.get("altitude", "< 200m"),
        "AboveGroundArea": project_info.get("floor_area", 0),
        "BaseArea": project_info.get("base_area", 0),
        "NewConstructionArea": project_info.get("building_area", 0),
        "TotalFloorArea": project_info.get("floor_area", 0),
        "ConstructionRate": project_info.get("coverage_rate", 0),
        "VolumeRate": 0,
        "ApplyWholeBase": False, "ApplyPartialBase": False, "PartialBaseArea": 0,
        "RealConstructionRate": 0, "RealVolumeRate": 0, "UndergroundBuilding": False,
        "IsPublic": proj.get("is_public", 1),
        "IsNoCheckGreenBuildMaterial": 1, "IsHillsideBuilding": -1, "IsGroundwaterLess1m": -1,
        "ApplicationGridData": [{
            "Catalog": use_type,
            "Description": "",
            "Type": "空調型建築",
            "FloorArea": project_info.get("floor_area", 0),
            "IsOpenHalf": False,
            "EnergySaving": make_energy_saving(),
        }],
        "DifficultyArea": 0,
        "MinimumGreenArea": project_info.get("base_area", 0) * 0.5,
        "CarbonBenchmark": site.get("carbon_benchmark", "前二類以外之建築基地"),
        "Alpha08": False,
        "PlantGridData": [],
        "IsDrillingReport": -1,
        "SoilPermeability": site.get("soil_permeability", 1e-5),
        "SoilClassification": site.get("soil_classification", "回填土"),
        "BaseInfiltration": site.get("soil_permeability", 1e-5),
        "SchoolCampusOverallAssessment": False,
        "Area_A1": 0, "Area_A2": 0, "PermeableThicknessArea_A2": 0,
        "StructureType_A2": 0, "PermeableThicknessStructureType_A2": 0,
        "Volume_V3": 0, "Area_A4": 0, "Volume_V4": 0,
        "StorageType": "組合式蓄水框架",
        "Area_A5": 0, "Volume_V5": 0,
        "InfiltrationPipeLength": project_info.get("infiltration_length", 0),
        "HoleRate": 0,
        "InfiltrationWells": project_info.get("infiltration_wells", 0),
        "InfiltrationMaterial": "透水磚或透水混凝土",
        "InfiltrationSideLength": 0,
        "CalcType": 0,
        "VacmType": 0, "Vacm": 1, "VacmSAk": 0, "VacmWindowData": [],
        "VacType": 0, "Vac": 1, "VacSAk": 0, "VacWindowData": [],
        "ApplicationNumber": "", "ApplicationDate": "", "ApplicantsName": "",
        "Address": "", "UseClassGroup": use_type,
        "OriginalBuildingLicenseNumber": "",
        "BaseUseArea": 0, "ApplyForFloorArea": 0, "TotalOutdoorFloorAreaOfTheBuilding": 0,
        "NoCheckGreenBuildMaterialInside": False, "NoCheckGreenBuildMaterialOutside": False,
        "GreenBuildMaterialA1Data": [], "GreenBuildMaterialA2Data": [],
        "GreenBuildMaterialA3Data": [], "GreenBuildMaterialA4Data": [],
        "GreenBuildMaterialA5Data": [],
        "GreenBuildMaterialGi1Data": [], "GreenBuildMaterialGi2Data": [],
        "GreenBuildMaterialGi3Data": [], "GreenBuildMaterialGi4Data": [],
        "GreenBuildMaterialGi5Data": [], "GreenBuildMaterialG2Data": [],
        "NoReviewRainwaterStorage": False, "Use2022RainStorageRule": False,
        "RainfallStation": "", "DailyRainfallAverage": 0, "DailyRainfallProbability": 0,
        "WaterStorage": 0, "RainStorageCity": "", "RainStorageTown": "",
        "AverageDailyRainfall": 0, "RecommendedStorageDays": 0,
        "RainCollectionArea": 0, "ActualRainwaterCapacity": 0,
        "RainwaterUtilizationAmount": 0, "WaterConsumptionData": [],
        "UseRecyclingWasteWater": False, "RegenerationCapacity": 0, "IsWaterQualified": -1,
        "Version": "1.5.1",
        "DataPlantUser": dict(DATA_PLANT_USER),
        "DataHeatBuildMaterialUser": {}, "DataHeatGlassUser": {}, "DataStructureUser": {},
    }

    # Build window base entries
    for wno, count in sorted(window_counts.items()):
        if wno.startswith("D"):
            continue
        w, h = 1.0, 1.0
        or_ = win_cfg.get("open_ratio", 0.5)
        oww = round(w * or_, 2)
        gbf["DataWindowBaseUser"][wno] = {
            "WindowNo": wno, "WindowWidth": w, "WindowHeight": h,
            "Agsi": round(w * h, 4), "OWW": oww, "OWH": h, "OWA": round(oww * h, 4),
            "Catalog": win_cfg.get("glass", "單層透明玻璃"),
            "Color": win_cfg.get("color", "平板玻璃"),
            "Code": win_cfg.get("glass_code", "P5"),
            "GlassName": win_cfg.get("glass_name", "單層玻璃"),
            "Thickness": win_cfg.get("thickness", "3"),
            "Frame": win_cfg.get("frame", "鋁門窗窗框"),
            "Type": 1, "IsGreenBuildMaterial": False,
            "ConstructionCode": "", "MaterialName": "",
            "ApprovedDocumentNumber": "", "DataCheck": "",
        }

    # Build plant entries
    for p in plants:
        cat_info = DATA_PLANT_USER.get(p["catalog"], {})
        gbf["PlantGridData"].append({
            "Catalog": p["catalog"], "Name": p["name"],
            "Gi": cat_info.get("Gi", 0.3),
            "Depth": cat_info.get("Depth", 0), "DepthOther": cat_info.get("DepthOther", 0),
            "Area": cat_info.get("Area", 0), "Area2": cat_info.get("Area2", 0),
            "HasNumbers": False,
            "IsNative": False, "IsExotic": True, "IsLure": False,
            "AreaName": None, "DensePlanting": False,
            "PlantNumber": 0, "PlantingArea": p.get("area", 0),
            "PlantingAreaAi": p.get("area", 0),
            "CheckedDepth": False, "CheckedDepthOther": True,
            "CheckedDepthNoAssessment": False,
            "CheckedArea": True, "CheckedArea2": False,
            "CheckedAreaNoAssessment": False,
            "CheckedOver2Native": True, "CheckedOneNative": False,
            "Status": 1,
        })

    return gbf


def write_gbf(data, path):
    with open(path, "w", encoding="utf-8-sig") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))


# ─────────────────────────────────────────────────────────
# Custom tkinter widgets
# ─────────────────────────────────────────────────────────

class StyledButton(tk.Frame):
    """Retro-styled button with thin black border + amber hover."""

    def __init__(self, parent, text="", command=None,
                 bg_color=C["bg"], fg_color=C["text"], hover_color=C["accent"],
                 font=FONT_BTN, padx=24, pady=10, **kw):
        super().__init__(parent, bg=C["border"], padx=1, pady=1)
        self._label = tk.Label(self, text=text, font=font,
                               fg=fg_color, bg=bg_color,
                               padx=padx, pady=pady, cursor="hand2")
        self._label.pack(fill="both", expand=True)
        self._cmd = command
        self._bg = bg_color
        self._fg = fg_color
        self._hover_bg = hover_color
        self._hover_fg = C["text"]
        self._enabled = True
        self._label.bind("<Enter>", self._on_enter)
        self._label.bind("<Leave>", self._on_leave)
        self._label.bind("<Button-1>", self._on_click)

    def _on_enter(self, e):
        if self._enabled:
            self._label.configure(bg=self._hover_bg, fg=self._hover_fg)

    def _on_leave(self, e):
        if self._enabled:
            self._label.configure(bg=self._bg, fg=self._fg)

    def _on_click(self, event=None):
        if self._enabled and self._cmd:
            self._cmd()

    def configure_state(self, enabled, text=None):
        self._enabled = enabled
        if text:
            self._label.configure(text=text)
        if enabled:
            self._label.configure(bg=self._bg, fg=self._fg, cursor="hand2")
            self.configure(bg=C["border"])
        else:
            self._label.configure(bg=C["surface0"], fg=C["overlay"], cursor="arrow")
            self.configure(bg=C["surface2"])


class LabeledEntry(tk.Frame):
    """Retro-styled label + entry with thin border."""

    def __init__(self, parent, label, default="", width=40, **kw):
        super().__init__(parent, bg=C["bg"])
        tk.Label(self, text=label, font=FONT_SMALL, fg=C["subtext"], bg=C["bg"],
                 anchor="w", width=16).pack(side="left", padx=(0, 8))
        self.var = tk.StringVar(value=default)
        # Border frame
        border = tk.Frame(self, bg=C["border"], padx=1, pady=1)
        border.pack(side="left", fill="x", expand=True)
        self.entry = tk.Entry(border, textvariable=self.var, width=width,
                              font=FONT_BODY, fg=C["text"], bg="#FFFFFF",
                              insertbackground=C["text"], relief="flat",
                              highlightthickness=0)
        self.entry.pack(fill="x", expand=True, ipady=4, padx=1, pady=1)

    def get(self):
        return self.var.get().strip()

    def set(self, val):
        self.var.set(val)


class LabeledCombo(tk.Frame):
    """Retro-styled label + dropdown with thin border."""

    def __init__(self, parent, label, options, default="", width=38, **kw):
        super().__init__(parent, bg=C["bg"])
        tk.Label(self, text=label, font=FONT_SMALL, fg=C["subtext"], bg=C["bg"],
                 anchor="w", width=16).pack(side="left", padx=(0, 8))
        self.var = tk.StringVar(value=default)
        # Border frame
        border = tk.Frame(self, bg=C["border"], padx=1, pady=1)
        border.pack(side="left", fill="x", expand=True)

        # Style the combobox
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Retro.TCombobox",
                         fieldbackground="#FFFFFF",
                         background=C["bg"],
                         foreground=C["text"],
                         borderwidth=0,
                         arrowcolor=C["text"],
                         selectbackground=C["accent2"],
                         selectforeground=C["text"])
        style.map("Retro.TCombobox",
                   fieldbackground=[("readonly", "#FFFFFF")],
                   selectbackground=[("readonly", C["accent2"])],
                   selectforeground=[("readonly", C["text"])])

        self.combo = ttk.Combobox(border, textvariable=self.var,
                                   values=options, width=width,
                                   font=FONT_BODY, state="readonly",
                                   style="Retro.TCombobox")
        self.combo.pack(fill="x", expand=True, ipady=4, padx=1, pady=1)

        # Set default if in options
        if default in options:
            self.combo.set(default)
        elif options:
            self.combo.set(options[0])

    def get(self):
        return self.var.get().strip()

    def set(self, val):
        self.var.set(val)


# ─────────────────────────────────────────────────────────
# Main Application
# ─────────────────────────────────────────────────────────

class GreenBuildingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("綠建築 GBF 產生器")
        self.geometry("860x680")
        self.configure(bg=C["bg"])
        self.minsize(780, 600)

        # Set window icon (title bar + taskbar)
        self._set_app_icon()

        self.dwg_path = None
        self.output_path = None
        self.cfg = load_config()
        self._has_saved_config = os.path.exists(CONFIG_PATH)

        self._build_ui()

        # Auto-save settings on window close
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Enable drag-and-drop (must be after window is fully built)
        if HAS_DND:
            self.update_idletasks()  # ensure window handle exists
            windnd.hook_dropfiles(self, func=self._on_drop_files, force_unicode=True)

        # Show startup disclaimer
        self.after(200, self._show_disclaimer)

        # Check if first run (no authorization)
        if not self.cfg.get("office", {}).get("authorization"):
            self.after(300, lambda: self._show_tab("settings"))

    def _set_app_icon(self):
        """Set window icon for title bar and taskbar."""
        # On Windows: set AppUserModelID so taskbar shows our icon, not Python's
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "GreenBuildingGBF.v1.5.1")
        except Exception:
            pass

        # Look for app.ico: next to EXE, or in PyInstaller temp dir
        ico_path = os.path.join(APP_DIR, "app.ico")
        if not os.path.exists(ico_path) and getattr(sys, "_MEIPASS", None):
            ico_path = os.path.join(sys._MEIPASS, "app.ico")
        if os.path.exists(ico_path):
            try:
                self.iconbitmap(ico_path)
                return
            except Exception:
                pass

        # Fallback: generate icon in memory from PIL if available
        try:
            from PIL import Image, ImageDraw, ImageFont, ImageTk

            def _make_icon(size):
                img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
                d = ImageDraw.Draw(img)
                m = max(1, size // 16)
                r = max(2, size // 8)
                d.rounded_rectangle([m, m, size-m-1, size-m-1], radius=r,
                                     fill="#F5F3EE", outline="#1A1A1A",
                                     width=max(1, size // 64))
                bar_h = max(2, size // 8)
                d.rectangle([m+1, m+1, size-m-2, m+bar_h], fill="#F0A500")
                try:
                    fs = max(8, size * 38 // 100)
                    font = ImageFont.truetype("consola.ttf", fs)
                except Exception:
                    font = ImageFont.load_default()
                bbox = d.textbbox((0, 0), "GBF", font=font)
                tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
                tx = (size - tw) // 2
                ty = m + bar_h + (size - m*2 - bar_h - th) // 2 - max(1, size//20)
                d.text((tx, ty), "GBF", fill="#1A1A1A", font=font)
                ls = max(3, size // 6)
                lx = size - m - ls - max(2, size // 10)
                ly = size - m - ls - max(2, size // 12)
                d.ellipse([lx, ly, lx+ls, ly+ls], fill="#2D8A4E")
                return img

            # Set as window icon via PhotoImage
            icon_img = _make_icon(64)
            self._icon_photo = ImageTk.PhotoImage(icon_img)
            self.iconphoto(True, self._icon_photo)
        except Exception:
            pass

    # ── Startup disclaimer ───────────────────────────

    def _show_disclaimer(self):
        """Show a startup disclaimer dialog."""
        dlg = tk.Toplevel(self)
        dlg.title("使用須知")
        dlg.configure(bg=C["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.transient(self)

        # Try to set same icon as main window
        ico_path = os.path.join(APP_DIR, "app.ico")
        if not os.path.exists(ico_path) and getattr(sys, "_MEIPASS", None):
            ico_path = os.path.join(sys._MEIPASS, "app.ico")
        if os.path.exists(ico_path):
            try:
                dlg.iconbitmap(ico_path)
            except Exception:
                pass

        # ── Content frame with border ──
        border = tk.Frame(dlg, bg=C["border"], padx=1, pady=1)
        border.pack(padx=20, pady=20)

        inner = tk.Frame(border, bg=C["bg"], padx=28, pady=24)
        inner.pack()

        # Amber accent bar
        bar = tk.Frame(inner, bg=C["accent"], height=3)
        bar.pack(fill="x", pady=(0, 16))

        # Warning icon + title
        tk.Label(inner, text="⚠  使用須知", font=FONT_HEADING,
                 fg=C["text"], bg=C["bg"]).pack(anchor="w")

        tk.Frame(inner, bg=C["bg"], height=12).pack()

        # Message body
        msg = (
            "本工具透過自動掃描 AutoCAD 圖面，擷取綠建築評估\n"
            "所需資料並產出 GBF 檔案，可大幅節省人工輸入時間。\n"
            "\n"
            "惟自動擷取之資料仍可能存在遺漏或誤判，\n"
            "匯入綠建築評估系統後，請務必逐項核對確認。"
        )
        tk.Label(inner, text=msg, font=FONT_BODY, fg=C["subtext"],
                 bg=C["bg"], justify="left").pack(anchor="w")

        tk.Frame(inner, bg=C["bg"], height=18).pack()

        # OK button
        btn_frame = tk.Frame(inner, bg=C["bg"])
        btn_frame.pack()

        ok_border = tk.Frame(btn_frame, bg=C["border"], padx=1, pady=1)
        ok_border.pack()
        ok_btn = tk.Label(ok_border, text="  我已了解，開始使用  ", font=FONT_BTN,
                          fg=C["bg"], bg=C["accent"], cursor="hand2",
                          padx=16, pady=6)
        ok_btn.pack()
        ok_btn.bind("<Enter>", lambda e: ok_btn.configure(bg=C["accent2"]))
        ok_btn.bind("<Leave>", lambda e: ok_btn.configure(bg=C["accent"]))
        ok_btn.bind("<Button-1>", lambda e: dlg.destroy())

        # Center dialog on parent
        dlg.update_idletasks()
        w = dlg.winfo_width()
        h = dlg.winfo_height()
        px = self.winfo_x() + (self.winfo_width() - w) // 2
        py = self.winfo_y() + (self.winfo_height() - h) // 2
        dlg.geometry(f"+{px}+{py}")

    # ── Build UI ─────────────────────────────────────

    def _build_ui(self):
        # ── Top bar: thin black border bottom ──
        top = tk.Frame(self, bg=C["bg"], height=60)
        top.pack(fill="x")
        top.pack_propagate(False)

        # Title left
        title_frame = tk.Frame(top, bg=C["bg"])
        title_frame.pack(side="left", padx=20, pady=8)

        tk.Label(title_frame, text="GreenBuildingGBF", font=FONT_TITLE,
                 fg=C["text"], bg=C["bg"]).pack(side="left")
        tk.Label(title_frame, text="  v1.5.1", font=FONT_TAG,
                 fg=C["overlay"], bg=C["bg"]).pack(side="left", pady=(8, 0))

        # Tab buttons right
        self._tab_btns = {}
        tab_frame = tk.Frame(top, bg=C["bg"])
        tab_frame.pack(side="right", padx=16, pady=12)

        for name, label in [("settings", "設定"), ("scan", "掃描產出")]:
            outer = tk.Frame(tab_frame, bg=C["border"], padx=1, pady=1)
            outer.pack(side="right", padx=3)
            b = tk.Label(outer, text=f"  {label}  ", font=FONT_BODY,
                         fg=C["text"], bg=C["bg"], cursor="hand2",
                         padx=10, pady=4)
            b.pack()
            b.bind("<Button-1>", lambda e, n=name: self._show_tab(n))
            b.bind("<Enter>", lambda e, b=b: b.configure(bg=C["accent2"]))
            b.bind("<Leave>", lambda e, b=b, n=name: b.configure(
                bg=C["accent"] if self._active_tab == n else C["bg"]))
            self._tab_btns[name] = (b, outer)

        # Border line under top bar
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # Content area
        self._content = tk.Frame(self, bg=C["bg"])
        self._content.pack(fill="both", expand=True)

        # Build pages
        self._pages = {}
        self._build_scan_page()
        self._build_settings_page()

        self._active_tab = None
        self._show_tab("scan")

    def _show_tab(self, name):
        # Auto-save settings when leaving settings page
        if self._active_tab == "settings" and name != "settings":
            self._auto_save_settings()
        self._active_tab = name
        for n, page in self._pages.items():
            if n == name:
                page.pack(fill="both", expand=True)
            else:
                page.pack_forget()
        for n, (btn, outer) in self._tab_btns.items():
            if n == name:
                btn.configure(bg=C["accent"], fg=C["text"])
            else:
                btn.configure(bg=C["bg"], fg=C["text"])

    # ── Scan page ────────────────────────────────────

    def _build_scan_page(self):
        page = tk.Frame(self._content, bg=C["bg"])
        self._pages["scan"] = page

        # ── Hero section ──
        hero = tk.Frame(page, bg=C["bg"])
        hero.pack(pady=(24, 4))

        tk.Label(hero, text="DWG", font=("Consolas", 26, "bold"),
                 fg=C["text"], bg=C["bg"]).pack(side="left")
        tk.Label(hero, text="  -->  ", font=("Consolas", 20),
                 fg=C["accent"], bg=C["bg"]).pack(side="left")
        tk.Label(hero, text="GBF", font=("Consolas", 26, "bold"),
                 fg=C["text"], bg=C["bg"]).pack(side="left")

        tk.Label(page, text="從 AutoCAD 圖面自動掃描，產出綠建築評估系統檔案",
                 font=FONT_SMALL, fg=C["overlay"], bg=C["bg"]).pack(pady=(0, 12))

        # ── Drop zone / File picker — bordered card ──
        drop_border = tk.Frame(page, bg=C["border"], padx=1, pady=1)
        drop_border.pack(pady=(8, 12), padx=40, fill="x")

        self._drop_zone = tk.Frame(drop_border, bg="#FFFFFF", cursor="hand2")
        self._drop_zone.pack(fill="x")

        # Top row: INPUT tag + browse button
        top_row = tk.Frame(self._drop_zone, bg="#FFFFFF")
        top_row.pack(fill="x", padx=12, pady=(8, 0))

        tk.Label(top_row, text="INPUT", font=FONT_TAG,
                 fg=C["overlay"], bg="#FFFFFF").pack(side="left")

        browse_border = tk.Frame(top_row, bg=C["border"], padx=1, pady=1)
        browse_border.pack(side="right")
        browse_btn = tk.Label(browse_border, text=" 瀏覽... ", font=FONT_SMALL,
                              fg=C["text"], bg=C["bg"], cursor="hand2",
                              padx=8, pady=2)
        browse_btn.pack()
        browse_btn.bind("<Button-1>", self._select_file)
        browse_btn.bind("<Enter>", lambda e: browse_btn.configure(bg=C["accent2"]))
        browse_btn.bind("<Leave>", lambda e: browse_btn.configure(bg=C["bg"]))

        # Center: drop area with dashed border feel
        drop_center = tk.Frame(self._drop_zone, bg="#FFFFFF")
        drop_center.pack(fill="x", padx=12, pady=(6, 10))

        # Dashed border via a label inside a colored frame
        self._drop_inner = tk.Label(drop_center,
            text="將 DWG 檔案拖曳至此\n或點擊選擇檔案",
            font=FONT_BODY, fg=C["overlay"], bg=C["bg2"],
            padx=20, pady=16, cursor="hand2",
            relief="groove", bd=1)
        self._drop_inner.pack(fill="x")
        self._drop_inner.bind("<Button-1>", self._select_file)

        # File name display
        self.file_var = tk.StringVar(value="")
        self.file_label = tk.Label(self._drop_zone, textvariable=self.file_var,
                                   font=("Microsoft JhengHei UI", 9, "bold"),
                                   fg=C["green"], bg="#FFFFFF", anchor="w")
        self.file_label.pack(fill="x", padx=14, pady=(0, 6))

        # Also make the whole drop zone clickable
        for widget in [self._drop_zone, drop_center]:
            widget.bind("<Button-1>", self._select_file)

        # ── Output folder selector ──
        out_border = tk.Frame(page, bg=C["border"], padx=1, pady=1)
        out_border.pack(pady=(0, 8), padx=40, fill="x")

        out_inner = tk.Frame(out_border, bg="#FFFFFF")
        out_inner.pack(fill="x")

        out_row = tk.Frame(out_inner, bg="#FFFFFF")
        out_row.pack(fill="x", padx=12, pady=8)

        tk.Label(out_row, text="OUTPUT", font=FONT_TAG,
                 fg=C["overlay"], bg="#FFFFFF").pack(side="left", padx=(0, 10))

        self.outdir_var = tk.StringVar(value="(預設: 與 DWG 同資料夾)")
        tk.Label(out_row, textvariable=self.outdir_var,
                 font=FONT_SMALL, fg=C["subtext"], bg="#FFFFFF",
                 anchor="w").pack(side="left", fill="x", expand=True)

        outdir_btn_border = tk.Frame(out_row, bg=C["border"], padx=1, pady=1)
        outdir_btn_border.pack(side="right")
        outdir_btn = tk.Label(outdir_btn_border, text=" 變更... ", font=FONT_SMALL,
                              fg=C["text"], bg=C["bg"], cursor="hand2",
                              padx=8, pady=2)
        outdir_btn.pack()
        outdir_btn.bind("<Button-1>", self._select_outdir)
        outdir_btn.bind("<Enter>", lambda e: outdir_btn.configure(bg=C["accent2"]))
        outdir_btn.bind("<Leave>", lambda e: outdir_btn.configure(bg=C["bg"]))

        reset_btn_border = tk.Frame(out_row, bg=C["border"], padx=1, pady=1)
        reset_btn_border.pack(side="right", padx=(0, 4))
        reset_btn = tk.Label(reset_btn_border, text=" 重設 ", font=FONT_SMALL,
                             fg=C["text"], bg=C["bg"], cursor="hand2",
                             padx=6, pady=2)
        reset_btn.pack()
        reset_btn.bind("<Button-1>", self._reset_outdir)
        reset_btn.bind("<Enter>", lambda e: reset_btn.configure(bg=C["accent2"]))
        reset_btn.bind("<Leave>", lambda e: reset_btn.configure(bg=C["bg"]))

        self.custom_outdir = None  # None = use default logic

        # ── Action buttons ──
        btn_row = tk.Frame(page, bg=C["bg"])
        btn_row.pack(pady=8)

        self.scan_btn = StyledButton(btn_row, text="  掃描並產出 GBF  ",
                                       command=self._start_scan,
                                       bg_color=C["accent"], hover_color=C["accent2"],
                                       fg_color=C["text"])
        self.scan_btn.pack(side="left", padx=6)

        self.open_btn = StyledButton(btn_row, text="  開啟輸出資料夾  ",
                                       command=self._open_output,
                                       bg_color=C["bg"], hover_color=C["surface0"],
                                       fg_color=C["text"])
        self.open_btn.pack(side="left", padx=6)
        self.open_btn.configure_state(False)

        # ── Log area — terminal style ──
        log_outer = tk.Frame(page, bg=C["border"], padx=1, pady=1)
        log_outer.pack(padx=30, pady=(8, 16), fill="both", expand=True)

        # Log header bar
        log_header = tk.Frame(log_outer, bg=C["bg2"], height=24)
        log_header.pack(fill="x")
        log_header.pack_propagate(False)
        tk.Label(log_header, text="  TERMINAL OUTPUT", font=FONT_TAG,
                 fg=C["overlay"], bg=C["bg2"], anchor="w").pack(side="left", padx=8)
        # Decorative dots
        for color in [C["red"], C["accent"], C["green"]]:
            dot = tk.Frame(log_header, bg=color, width=8, height=8)
            dot.pack(side="right", padx=2, pady=8)

        self.log = scrolledtext.ScrolledText(
            log_outer, font=FONT_MONO,
            fg=C["text"], bg="#FFFFFF",
            insertbackground=C["text"],
            relief="flat", state="disabled",
            selectbackground=C["accent2"],
            wrap="word")
        self.log.pack(fill="both", expand=True, padx=0, pady=0)

        # Tag configs for colored log
        self.log.tag_configure("info", foreground=C["subtext"])
        self.log.tag_configure("good", foreground=C["green"])
        self.log.tag_configure("warn", foreground=C["peach"])
        self.log.tag_configure("err", foreground=C["red"])
        self.log.tag_configure("head", foreground=C["text"], font=(FONT_MONO[0], FONT_MONO[1], "bold"))

    # ── Settings page ────────────────────────────────

    def _build_settings_page(self):
        page = tk.Frame(self._content, bg=C["bg"])
        self._pages["settings"] = page

        # Scrollable
        canvas = tk.Canvas(page, bg=C["bg"], highlightthickness=0)
        scrollbar = tk.Scrollbar(page, orient="vertical", command=canvas.yview,
                                  bg=C["surface0"], troughcolor=C["bg2"],
                                  highlightthickness=0, bd=0)
        scroll_frame = tk.Frame(canvas, bg=C["bg"])

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=30, pady=16)
        scrollbar.pack(side="right", fill="y")

        # Mouse wheel
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # ── Office info section ──
        self._section_label(scroll_frame, "事務所資料",
                           "您的事務所基本資料與綠建築系統授權碼")

        office = self.cfg.get("office", {})
        self.se = {}  # settings entries

        fields = [
            ("name",           "事務所名稱",       office.get("name", "")),
            ("tel",            "電話",             office.get("tel", "")),
            ("address",        "地址",             office.get("address", "")),
            ("visa_holder",    "簽證建築師",        office.get("visa_holder", "")),
            ("certificate_no", "開業證書字號",      office.get("certificate_no", "")),
            ("authorization",  "授權碼 *",         office.get("authorization", "")),
            ("verification",   "驗證碼 *",         office.get("verification", "")),
            ("term_start",     "授權起始日",        office.get("term_start", "")),
            ("term_end",       "授權到期日",        office.get("term_end", "")),
        ]
        for key, label, val in fields:
            e = LabeledEntry(scroll_frame, label, val, width=50)
            e.pack(fill="x", padx=10, pady=3)
            self.se[f"office.{key}"] = e

        # Hint
        tk.Label(scroll_frame,
                 text="* 授權碼與驗證碼請從既有的 .GBF 檔案中複製",
                 font=FONT_CAPTION, fg=C["peach"], bg=C["bg"], anchor="w").pack(padx=12, pady=(2, 10))

        # ── Window defaults ──
        self._section_label(scroll_frame, "窗戶預設值", "預設玻璃與窗框設定")

        win_cfg = self.cfg.get("window_defaults", {})

        # 玻璃類型 — dropdown
        GLASS_OPTIONS = [
            "單層透明玻璃", "雙層透明玻璃", "低輻射(Low-E)玻璃",
            "反射玻璃", "有色玻璃", "膠合玻璃",
        ]
        e = LabeledCombo(scroll_frame, "玻璃類型", GLASS_OPTIONS,
                          win_cfg.get("glass", "單層透明玻璃"))
        e.pack(fill="x", padx=10, pady=3)
        self.se["window.glass"] = e

        # 玻璃代碼 — dropdown
        GLASS_CODE_OPTIONS = [
            "P5", "P6", "P8", "PL6", "PH6", "RE6", "RE8",
            "RG6", "AB6", "AB8", "LA6",
        ]
        e = LabeledCombo(scroll_frame, "玻璃代碼", GLASS_CODE_OPTIONS,
                          win_cfg.get("glass_code", "P5"))
        e.pack(fill="x", padx=10, pady=3)
        self.se["window.glass_code"] = e

        # 厚度 — dropdown
        THICKNESS_OPTIONS = ["3", "5", "6", "8", "10", "12"]
        e = LabeledCombo(scroll_frame, "厚度 (mm)", THICKNESS_OPTIONS,
                          win_cfg.get("thickness", "3"))
        e.pack(fill="x", padx=10, pady=3)
        self.se["window.thickness"] = e

        # 窗框類型 — dropdown
        FRAME_OPTIONS = [
            "鋁門窗窗框", "塑鋼窗框", "木框", "鋁包木框", "不鏽鋼窗框",
        ]
        e = LabeledCombo(scroll_frame, "窗框類型", FRAME_OPTIONS,
                          win_cfg.get("frame", "鋁門窗窗框"))
        e.pack(fill="x", padx=10, pady=3)
        self.se["window.frame"] = e

        # 可開窗比例 — free entry (numeric)
        e = LabeledEntry(scroll_frame, "可開窗比例",
                          str(win_cfg.get("open_ratio", 0.5)), width=50)
        e.pack(fill="x", padx=10, pady=3)
        self.se["window.open_ratio"] = e

        # ── Site defaults ──
        self._section_label(scroll_frame, "基地預設值", "基地與土壤相關參數")

        site = self.cfg.get("site_defaults", {})

        # 海拔高度 — dropdown
        ALTITUDE_OPTIONS = ["< 200m", "200-800m", "> 800m"]
        e = LabeledCombo(scroll_frame, "海拔高度", ALTITUDE_OPTIONS,
                          site.get("altitude", "< 200m"))
        e.pack(fill="x", padx=10, pady=3)
        self.se["site.altitude"] = e

        # 土壤分類 — dropdown
        SOIL_OPTIONS = ["回填土", "砂質土", "粘土", "礫石層", "岩盤"]
        e = LabeledCombo(scroll_frame, "土壤分類", SOIL_OPTIONS,
                          site.get("soil_classification", "回填土"))
        e.pack(fill="x", padx=10, pady=3)
        self.se["site.soil_classification"] = e

        # 土壤滲透係數 — free entry (numeric)
        e = LabeledEntry(scroll_frame, "土壤滲透係數",
                          str(site.get("soil_permeability", 1e-5)), width=50)
        e.pack(fill="x", padx=10, pady=3)
        self.se["site.soil_permeability"] = e

        # 碳足跡基準 — dropdown
        CARBON_OPTIONS = [
            "學校用地、公園用地",
            "前二類以外之建築基地",
        ]
        e = LabeledCombo(scroll_frame, "碳足跡基準", CARBON_OPTIONS,
                          site.get("carbon_benchmark", "前二類以外之建築基地"))
        e.pack(fill="x", padx=10, pady=3)
        self.se["site.carbon_benchmark"] = e

        # ── Project defaults ──
        self._section_label(scroll_frame, "案件預設值", "申請類型與用地分區")

        proj = self.cfg.get("project_defaults", {})

        # 申請類型 — dropdown
        APPLY_OPTIONS = [
            "建造執照申請（含變更設計）",
            "使用執照申請",
            "雜項執照申請",
        ]
        e = LabeledCombo(scroll_frame, "申請類型", APPLY_OPTIONS,
                          proj.get("apply_type", "建造執照申請（含變更設計）"))
        e.pack(fill="x", padx=10, pady=3)
        self.se["project.apply_type"] = e

        # 用地分區 — free entry (too many possibilities)
        e = LabeledEntry(scroll_frame, "用地分區",
                          proj.get("land_use_2", ""), width=50)
        e.pack(fill="x", padx=10, pady=3)
        self.se["project.land_use_2"] = e

        # 是否公有 — dropdown
        PUBLIC_OPTIONS = ["非公有建築", "公有建築"]
        public_val = proj.get("is_public", 1)
        public_display = "公有建築" if public_val == 0 else "非公有建築"
        e = LabeledCombo(scroll_frame, "是否公有", PUBLIC_OPTIONS, public_display)
        e.pack(fill="x", padx=10, pady=3)
        self.se["project.is_public"] = e

        # Save button
        tk.Frame(scroll_frame, bg=C["bg"], height=15).pack()
        save_row = tk.Frame(scroll_frame, bg=C["bg"])
        save_row.pack(pady=(5, 25))

        StyledButton(save_row, text="  儲存設定  ", command=self._save_settings,
                      bg_color=C["accent"], hover_color=C["accent2"]).pack(side="left", padx=8)

        self._save_status = tk.Label(save_row, text="", font=FONT_SMALL,
                                      fg=C["green"], bg=C["bg"])
        self._save_status.pack(side="left", padx=8)

        # Auto-save hint
        hint_text = ("💡 設定會在切換分頁或關閉程式時自動儲存"
                     if not self._has_saved_config
                     else "✓ 已載入上次儲存的設定（自動記憶）")
        hint_color = C["overlay"] if not self._has_saved_config else C["green"]
        tk.Label(scroll_frame, text=hint_text, font=FONT_CAPTION,
                 fg=hint_color, bg=C["bg"], anchor="w").pack(padx=12, pady=(0, 20))

    def _section_label(self, parent, title, subtitle=""):
        frame = tk.Frame(parent, bg=C["bg"])
        frame.pack(fill="x", padx=10, pady=(20, 6))
        # Section header with amber accent bar
        header_row = tk.Frame(frame, bg=C["bg"])
        header_row.pack(fill="x")
        tk.Frame(header_row, bg=C["accent"], width=4, height=18).pack(side="left", padx=(0, 8))
        tk.Label(header_row, text=title.upper(), font=FONT_HEADING,
                 fg=C["text"], bg=C["bg"], anchor="w").pack(side="left")
        if subtitle:
            tk.Label(frame, text=subtitle, font=FONT_CAPTION,
                     fg=C["overlay"], bg=C["bg"], anchor="w").pack(anchor="w", padx=12, pady=(2, 0))
        tk.Frame(frame, bg=C["border"], height=1).pack(fill="x", pady=(6, 0))

    def _save_settings(self):
        cfg = dict(DEFAULT_CONFIG)
        # Deep copy nested dicts
        for k in cfg:
            if isinstance(cfg[k], dict):
                cfg[k] = dict(cfg[k])

        # Read all entries
        for key, entry in self.se.items():
            section, field = key.split(".", 1)
            section_map = {
                "office": "office", "window": "window_defaults",
                "site": "site_defaults", "project": "project_defaults"
            }
            cfg_section = section_map.get(section, section)
            val = entry.get()
            # Convert is_public dropdown to int
            if field == "is_public":
                val = 0 if val == "公有建築" else 1
            # Try numeric conversion
            elif field in ("open_ratio", "soil_permeability"):
                try:
                    val = float(val)
                except ValueError:
                    pass
            cfg[cfg_section][field] = val

        save_config(cfg)
        self.cfg = cfg
        self._save_status.configure(text="已儲存!", fg=C["green"])
        self.after(2000, lambda: self._save_status.configure(text=""))

    def _auto_save_settings(self):
        """Silently save settings (called on tab switch and window close)."""
        try:
            if not hasattr(self, "se") or not self.se:
                return
            cfg = dict(DEFAULT_CONFIG)
            for k in cfg:
                if isinstance(cfg[k], dict):
                    cfg[k] = dict(cfg[k])
            for key, entry in self.se.items():
                section, field = key.split(".", 1)
                section_map = {
                    "office": "office", "window": "window_defaults",
                    "site": "site_defaults", "project": "project_defaults"
                }
                cfg_section = section_map.get(section, section)
                val = entry.get()
                if field == "is_public":
                    val = 0 if val == "公有建築" else 1
                elif field in ("open_ratio", "soil_permeability"):
                    try:
                        val = float(val)
                    except ValueError:
                        pass
                cfg[cfg_section][field] = val
            save_config(cfg)
            self.cfg = cfg
        except Exception:
            pass

    def _on_close(self):
        """Auto-save settings before closing the window."""
        self._auto_save_settings()
        self.destroy()

    # ── Scan actions ─────────────────────────────────

    def _select_file(self, event=None):
        path = filedialog.askopenfilename(
            title="選擇 DWG 檔案",
            filetypes=[("AutoCAD 圖檔", "*.dwg"), ("所有檔案", "*.*")]
        )
        if path:
            self._set_dwg(path)

    def _on_drop_files(self, file_list):
        """Handle files dropped onto the window."""
        for f in file_list:
            # force_unicode=True → str, but handle bytes just in case
            if isinstance(f, bytes):
                try:
                    path = f.decode("utf-8")
                except UnicodeDecodeError:
                    path = f.decode("cp950", errors="replace")
            else:
                path = str(f)
            path = path.strip().strip('"')
            if path.lower().endswith(".dwg"):
                self._set_dwg(path)
                return
        # No DWG found in dropped files
        self.file_var.set("(請拖入 .dwg 檔案)")

    def _set_dwg(self, path):
        """Set the DWG file path and update UI."""
        self.dwg_path = path
        name = os.path.basename(path)
        self.file_var.set(name)
        # Update drop zone visual
        self._drop_inner.configure(
            text=f"{name}\n已選取，按下方按鈕開始掃描",
            fg=C["text"], bg=C["surface0"])
        # If no custom output dir, show DWG's folder as default
        if not self.custom_outdir:
            self.outdir_var.set(os.path.dirname(path))

    def _select_outdir(self, event=None):
        """Let user choose a custom output folder."""
        folder = filedialog.askdirectory(title="選擇 GBF 輸出資料夾")
        if folder:
            self.custom_outdir = folder
            self.outdir_var.set(folder)

    def _reset_outdir(self, event=None):
        """Reset output folder to default logic."""
        self.custom_outdir = None
        if self.dwg_path:
            self.outdir_var.set(os.path.dirname(self.dwg_path))
        else:
            self.outdir_var.set("(預設: 與 DWG 同資料夾)")

    def _log(self, msg, tag="info"):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def _start_scan(self):
        # Check if credentials are set
        auth = self.cfg.get("office", {}).get("authorization", "")
        if not auth:
            if not messagebox.askyesno(
                "尚未設定授權碼",
                "授權碼目前為空。\n\n"
                "產出的 GBF 將不包含授權資訊。\n"
                "您可在「設定」頁面中填入授權碼。\n\n"
                "是否仍要繼續?"):
                return

        self.scan_btn.configure_state(False, "  掃描中...  ")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        threading.Thread(target=self._run_scan, daemon=True).start()

    def _run_scan(self):
        try:
            dump_path = connect_and_dump(self.dwg_path, lambda m: self.after(0, lambda: self._log(m)))
            self.after(0, lambda: self._log("解析資料中...", "info"))
            blocks, texts = parse_dump(dump_path)
            self.after(0, lambda: self._log(f"  {len(blocks)} 個圖塊, {len(texts)} 個文字", "info"))

            info = extract_project_info(blocks, texts)
            wins = extract_windows(blocks, texts)
            plants = extract_plants(texts)

            self.after(0, lambda: self._log(""))
            self.after(0, lambda: self._log("=== 掃描結果 ===", "head"))
            for k, v in info.items():
                if k != "drawings":
                    self.after(0, lambda k=k, v=v: self._log(f"  {k}: {v}", "info"))

            self.after(0, lambda: self._log(f"\n  窗戶: {len(wins)} 種", "info"))
            for w, c in sorted(wins.items()):
                self.after(0, lambda w=w, c=c: self._log(f"    {w}: {c} 個", "info"))

            self.after(0, lambda: self._log(f"\n  植栽: {len(plants)} 項", "info"))
            for p in plants:
                self.after(0, lambda p=p: self._log(
                    f"    {p['name']} ({p['catalog']}) {p['area']} m2", "info"))

            self.after(0, lambda: self._log("\n組裝 GBF 檔案中...", "info"))

            # Reload config in case user changed settings
            self.cfg = load_config()
            gbf = build_gbf(info, wins, plants, self.cfg)

            name = info.get("address", "green_building")
            # Sanitize filename
            name = re.sub(r'[\\/:*?"<>|]', '_', name)
            if self.custom_outdir:
                out_dir = self.custom_outdir
            elif self.dwg_path:
                out_dir = os.path.dirname(self.dwg_path)
            else:
                out_dir = APP_DIR
            self.output_path = os.path.join(out_dir, f"{name}.GBF")
            write_gbf(gbf, self.output_path)

            self.after(0, lambda: self._log(""))
            self.after(0, lambda: self._log("=== 完成 ===", "head"))
            self.after(0, lambda: self._log(f"  輸出檔案: {self.output_path}", "good"))
            self.after(0, lambda: self._log(f"  用途類型: {gbf['UseClassGroup']}", "good"))
            self.after(0, lambda: self._log(f"  基地面積: {gbf['BaseArea']} m2", "good"))
            self.after(0, lambda: self._log(f"  樓地板面積: {gbf['TotalFloorArea']} m2", "good"))

            self.after(0, lambda: self.open_btn.configure_state(True))

        except Exception as e:
            self.after(0, lambda: self._log(f"\n[ERROR] {e}", "err"))
            import traceback
            tb = traceback.format_exc()
            self.after(0, lambda: self._log(tb, "err"))

        self.after(0, lambda: self.scan_btn.configure_state(True, "  掃描並產出 GBF  "))

    def _open_output(self):
        if self.output_path and os.path.exists(self.output_path):
            os.startfile(os.path.dirname(self.output_path))


# ─────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = GreenBuildingApp()
    app.mainloop()
