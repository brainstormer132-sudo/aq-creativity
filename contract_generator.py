import json
import os
import re
import shutil
from datetime import datetime
from pathlib import Path

# =================================================
# CONFIGURATION
# =================================================
APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "contract_config.json"

PREFERRED_OUTPUT_DIR = r"C:\Users\siraj\OneDrive - AQ Creativity\CONTRACTS"
LEGACY_DEFAULT_OUTPUT_DIR = str(APP_DIR / "output")

DEFAULT_CONFIG = {
    "output_dir": PREFERRED_OUTPUT_DIR,
    "temp_convert_dir": str(APP_DIR / "temp"),
    "generated_excel_path": str(APP_DIR / "generated.xlsx"),
    "archive_dir": str(APP_DIR / "archive"),
    "brand_id_file": str(APP_DIR / "brand_ids.json"),
    "template_map": {
        "after_pay": str(APP_DIR / "templates" / "Contract template(2).docx"),
        "pre_pay": str(APP_DIR / "templates" / "advance payment contract .docx"),
        "savola": str(APP_DIR / "templates" / "Savola Contract  FULL.docx"),
        "pre_savola": str(APP_DIR / "templates" / "Savola Contract advance .docx"),
        "crispy": str(APP_DIR / "templates" / "UGC Crispy Contract  .docx"),
        "santia": str(APP_DIR / "templates" / "Santia Contract  (1).docx"),
        "free_lancer": str(APP_DIR / "templates" / "Freelancer Contract .docx"),
    },
    "standard_template_key": "after_pay",
}

PLATFORM_MAP = {
    "instagram": "إنستقرام",
    "insta": "إنستقرام",
    "tiktok": "تيك توك",
    "tik tok": "تيك توك",
    "tik": "تيك توك",
    "snapchat": "سناب شات",
    "snap chat": "سناب شات",
    "snap": "سناب شات",
    "kick": "كيك",
    "youtube": "يوتيوب",
}

AD_TYPE_MAP = {
    "multi service": "خدمة متعددة",
    "home ad": "إعلان منزلي",
    "store visit": "إعلان زيارة",
}

ARABIC_DAYS = {
    "Monday": "الإثنين",
    "Tuesday": "الثلاثاء",
    "Wednesday": "الأربعاء",
    "Thursday": "الخميس",
    "Friday": "الجمعة",
    "Saturday": "السبت",
    "Sunday": "الأحد",
}

last_number = 0
_runtime = None
_dep_cache = {}
_brand_map_cache = None


def _require(dep_name: str, install_hint: str):
    cached = _dep_cache.get(dep_name)
    if cached is not None:
        return cached
    try:
        module = __import__(dep_name)
        _dep_cache[dep_name] = module
        return module
    except Exception as exc:
        raise RuntimeError(f"Missing dependency '{dep_name}'. Install with: {install_hint}") from exc


def load_config():
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, indent=2, ensure_ascii=False), encoding="utf-8")
        return DEFAULT_CONFIG.copy()

    raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    config = DEFAULT_CONFIG.copy()
    config.update({k: v for k, v in raw.items() if k in config and k != "template_map"})

    # Respect user-managed template list as source of truth.
    # If template_map exists in config file, use it directly.
    # Otherwise, fall back to defaults for first-run compatibility.
    template_map = raw.get("template_map")
    if isinstance(template_map, dict) and template_map:
        config["template_map"] = template_map
    else:
        config["template_map"] = DEFAULT_CONFIG["template_map"].copy()

    # keep standard key valid
    if config.get("standard_template_key") not in config["template_map"]:
        config["standard_template_key"] = next(iter(config["template_map"]))

    # Backward compatibility: if an old config still points to local ./output default,
    # move it to the requested AQ Creativity OneDrive contracts folder.
    if config.get("output_dir") in {"", LEGACY_DEFAULT_OUTPUT_DIR}:
        config["output_dir"] = PREFERRED_OUTPUT_DIR
        raw["output_dir"] = PREFERRED_OUTPUT_DIR
        CONFIG_PATH.write_text(json.dumps({**raw, "template_map": config["template_map"]}, indent=2, ensure_ascii=False), encoding="utf-8")

    return config


def _runtime_paths():
    global _runtime
    if _runtime is not None:
        return _runtime

    conf = load_config()
    output_dir = Path(conf["output_dir"])
    temp_convert_dir = Path(conf["temp_convert_dir"])
    generated_excel_path = Path(conf["generated_excel_path"])
    archive_dir = Path(conf["archive_dir"])
    brand_id_file = Path(conf["brand_id_file"])

    output_dir.mkdir(parents=True, exist_ok=True)
    temp_convert_dir.mkdir(parents=True, exist_ok=True)
    archive_dir.mkdir(parents=True, exist_ok=True)

    for f in temp_convert_dir.iterdir():
        if f.is_file():
            try:
                f.unlink()
            except Exception:
                pass

    template_map = conf["template_map"]
    standard_template = template_map.get(conf.get("standard_template_key", "after_pay"), template_map["after_pay"])

    _runtime = {
        "output_dir": output_dir,
        "temp_convert_dir": temp_convert_dir,
        "generated_excel_path": generated_excel_path,
        "archive_dir": archive_dir,
        "brand_id_file": brand_id_file,
        "template_map": template_map,
        "standard_template": standard_template,
    }
    return _runtime


def safe_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "", str(text)).strip()


def unique_path(path):
    base, ext = os.path.splitext(path)
    counter = 1
    new_path = path
    while os.path.exists(new_path):
        new_path = f"{base}_{counter}{ext}"
        counter += 1
    return new_path


def normalize_channel_name(raw):
    if not raw:
        return ""

    # Accept multiple separators used by UI payloads.
    chunks = [c.strip() for c in re.split(r"[|,\n]+", str(raw)) if c.strip()]
    parsed = []

    for chunk in chunks:
        if ":" in chunk:
            platform_raw, name_raw = chunk.split(":", 1)
            platform = platform_raw.strip()
            name = name_raw.strip().replace("@", "")
            parsed.append((platform, name))
        else:
            parsed.append(("", chunk.strip().replace("@", "")))

    # Requested behavior:
    # - one entry only -> @name
    # - two or more entries -> Platform: name per line (or name when no platform provided)
    if len(parsed) == 1:
        _, name = parsed[0]
        return f"@{name}" if name else ""

    lines = []
    for platform, name in parsed:
        if not name:
            continue
        if platform:
            lines.append(f"{platform}: {name}")
        else:
            lines.append(name)

    return "\n".join(lines)


def _get_brand_map():
    global _brand_map_cache
    if _brand_map_cache is not None:
        return _brand_map_cache

    runtime = _runtime_paths()
    brand_id_file = runtime["brand_id_file"]
    if brand_id_file.exists():
        _brand_map_cache = json.loads(brand_id_file.read_text(encoding="utf-8"))
    else:
        _brand_map_cache = {}
    return _brand_map_cache


def get_brand_id(brand_name):
    runtime = _runtime_paths()
    brand_id_file = runtime["brand_id_file"]

    brand_key = re.sub(r"[^A-Za-z0-9]", "", str(brand_name).upper())
    brand_map = _get_brand_map()

    if brand_key in brand_map:
        return brand_map[brand_key]

    new_id = str(max((int(v) for v in brand_map.values()), default=0) + 1)
    brand_map[brand_key] = new_id
    brand_id_file.write_text(json.dumps(brand_map, indent=4, ensure_ascii=False), encoding="utf-8")

    return new_id


def normalize_platform(raw):
    if not raw:
        return ""
    parts = re.split(r"[,+/&]+", raw.lower())
    result, seen = [], set()
    for p in parts:
        p = p.strip()
        if p in PLATFORM_MAP and p not in seen:
            result.append(PLATFORM_MAP[p])
            seen.add(p)
    return " و ".join(result)


def _load_last_number_from_generated_sheet():
    runtime = _runtime_paths()
    pd = _require("pandas", "pip install pandas openpyxl")
    generated_excel_path = runtime["generated_excel_path"]
    if not generated_excel_path.exists():
        return 0


    df_existing = pd.read_excel(generated_excel_path, dtype=str)

    id_column = None
    for candidate in ("ID", "id"):
        if candidate in df_existing.columns:
            id_column = candidate
            break

    if not id_column:
        return 0

    ids = df_existing[id_column].dropna().astype(str)
    nums = ids.str.extract(r"(\d{6})$")[0].dropna().astype(int)
    return int(nums.max()) if not nums.empty else 0


def _compose_amount_full(raw_amount: str) -> str:
    num2words = getattr(_require("num2words", "pip install num2words"), "num2words")

    clean = str(raw_amount).replace(",", "").strip()
    if not clean or clean == ".":
        raise ValueError("invalid amount")

    if "." in clean:
        r_str, h_str = clean.split(".", 1)
        h_str = (h_str + "00")[:2]
    else:
        r_str = clean
        h_str = "00"

    if not r_str.isdigit() or not h_str.isdigit():
        raise ValueError("non-numeric amount")

    r = int(r_str)
    h = int(h_str)

    amount_words = num2words(r, lang="ar") + " ريال سعودي"
    if h > 0:
        amount_words += " و " + num2words(h, lang="ar") + " هللة"

    amount_number = f"({r}.{h_str})" if h > 0 else f"({r})"
    rtl_start = "\u202B"
    rtl_end = "\u202C"
    return rtl_start + f"{amount_number} {amount_words}" + rtl_end


def _validate_template_path(template_path: str):
    if not Path(template_path).exists():
        raise FileNotFoundError(f"Template file not found: {template_path}")




def append_generated_rows(rows):
    if not rows:
        return
    runtime = _runtime_paths()
    pd = _require("pandas", "pip install pandas openpyxl")

    ordered_cols = [
        "id",
        "influencer_name_as_per_license",
        "license_number",
        "city_as_per_license",
        "neighbourhood_as_per_license",
        "brand_name",
        "platform",
        "channel_name",
        "ad_types",
        "amount",
        "bank_name",
        "account_name",
        "iban",
        "account_number",
        "swift_code",
        "contract_type",
        "date",
    ]

    output_rows = pd.DataFrame(rows).reindex(columns=ordered_cols)
    if runtime["generated_excel_path"].exists():
        existing = pd.read_excel(runtime["generated_excel_path"], dtype=str)
        output_rows = pd.concat([existing, output_rows], ignore_index=True)

    output_rows.to_excel(runtime["generated_excel_path"], index=False)



def batch_convert_docx_to_pdf(docx_paths, output_dir):
    if not docx_paths:
        return []

    runtime = _runtime_paths()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    docx2pdf_module = _require("docx2pdf", "pip install docx2pdf")
    convert = getattr(docx2pdf_module, "convert")

    # Reset temp conversion folder for a clean one-shot conversion pass.
    for item in runtime["temp_convert_dir"].iterdir():
        if item.is_file():
            try:
                item.unlink()
            except Exception:
                pass

    copied_stems = []
    for src in docx_paths:
        src_path = Path(src)
        if not src_path.exists():
            continue
        dst = runtime["temp_convert_dir"] / src_path.name
        if dst.exists():
            dst = runtime["temp_convert_dir"] / Path(unique_path(str(dst))).name
        shutil.copy(str(src_path), str(dst))
        copied_stems.append(dst.stem)

    if not copied_stems:
        return []

    convert(str(runtime["temp_convert_dir"]))

    generated = []
    for stem in copied_stems:
        temp_pdf = runtime["temp_convert_dir"] / f"{stem}.pdf"
        if not temp_pdf.exists():
            continue
        final_pdf = Path(unique_path(str(output_dir / f"{stem}.pdf")))
        os.replace(str(temp_pdf), str(final_pdf))
        generated.append(final_pdf)

    # Cleanup temp folder after conversion.
    for item in runtime["temp_convert_dir"].iterdir():
        if item.is_file():
            try:
                item.unlink()
            except Exception:
                pass

    return generated

def generate_contract_from_gui(context, append_excel: bool = True, return_metadata: bool = False, convert_pdf: bool = True):
    global last_number
    runtime = _runtime_paths()

    docxtpl_module = _require("docxtpl", "pip install docxtpl")

    docx_template_cls = getattr(docxtpl_module, "DocxTemplate")
    convert = None
    if convert_pdf:
        docx2pdf_module = _require("docx2pdf", "pip install docx2pdf")
        convert = getattr(docx2pdf_module, "convert")

    if last_number == 0:
        last_number = _load_last_number_from_generated_sheet()

    today = datetime.today()
    today_date = today.strftime("%d/%m/%Y")
    today_day_ar = ARABIC_DAYS.get(today.strftime("%A"), today.strftime("%A"))

    ctx = {k: str(v).strip() for k, v in (context or {}).items()}

    full_name = ctx.get("influencer_name_as_per_license", "")
    parts = full_name.split()
    ctx["name"] = full_name
    ctx["license_name"] = full_name
    ctx["name_2"] = f"{parts[0]} {parts[-1]}" if len(parts) >= 2 else full_name
    ctx["channel_name"] = normalize_channel_name(ctx.get("channel_name", ""))
    ctx["platform"] = normalize_platform(ctx.get("platform", ""))
    ctx["platform_smart"] = ctx["platform"]

    raw_ad = ctx.get("ad_types", "")
    if raw_ad:
        pieces = [p.strip() for p in raw_ad.split(",")]
        key = pieces[0].lower()
        qty = pieces[1] if len(pieces) > 1 and pieces[1].isdigit() else ""
        ad_ar = AD_TYPE_MAP.get(key, pieces[0])
        ctx["ad_types"] = f'{ad_ar} "{qty}"' if qty else ad_ar

    ctx["date"] = today_date
    ctx["day"] = f"({today_day_ar})"
    ctx["Amount_full"] = _compose_amount_full(ctx.get("amount", ""))

    brand_id = get_brand_id(ctx.get("brand_name", ""))
    last_number += 1
    contract_id = f"CTR{brand_id}{last_number:06d}"
    ctx["id"] = contract_id

    contract_key = re.sub(r"\s+", "_", ctx.get("contract_type", "").lower())
    template_path = runtime["template_map"].get(contract_key, runtime["standard_template"])
    _validate_template_path(template_path)

    doc = docx_template_cls(template_path)
    filename = safe_filename(contract_id)
    docx_path = unique_path(str(runtime["archive_dir"] / f"{filename}.docx"))
    doc.render(ctx)
    doc.save(docx_path)

    today_folder = runtime["output_dir"] / today.strftime("%Y-%m-%d")
    today_folder.mkdir(parents=True, exist_ok=True)

    pdf_path = ""
    if convert_pdf:
        # Fast path: convert only this DOCX directly to its final PDF target.
        pdf_path = unique_path(str(today_folder / f"{filename}.pdf"))
        try:
            convert(str(docx_path), str(pdf_path))
        except TypeError:
            # Backward compatibility for docx2pdf variants that may not support
            # two-argument conversion.
            temp_docx = runtime["temp_convert_dir"] / Path(docx_path).name
            shutil.copy(docx_path, str(temp_docx))
            convert(str(runtime["temp_convert_dir"]))
            for item in runtime["temp_convert_dir"].iterdir():
                if item.suffix.lower() == ".pdf":
                    os.replace(str(item), str(today_folder / item.name))
                elif item.is_file():
                    item.unlink()

    if append_excel:
        append_generated_rows([ctx])

    if return_metadata:
        return {
            "contract_id": contract_id,
            "pdf_path": str(pdf_path) if pdf_path else "",
            "docx_path": str(docx_path),
            "today_folder": str(today_folder),
            "row": ctx,
        }
    return contract_id
