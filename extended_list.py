from pathlib import Path
import re
import pandas as pd

BASE = Path(__file__).resolve().parent
OUTPUT = BASE / "Extended_Accommodation_List.xlsx"


# ---------- helpers ----------
def find_file(patterns):
    for pattern in patterns:
        matches = [
            f for f in sorted(BASE.glob(pattern))
            if not f.name.startswith("~$") and f.name != OUTPUT.name
        ]
        if matches:
            return matches[0]
    raise FileNotFoundError(f"No file found matching: {patterns}")


def norm(text):
    """lowercase, keep only letters/digits (handles stray spaces and | characters)"""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def norm_reg(text):
    return re.sub(r"[^A-Z0-9]", "", str(text).upper())


def norm_name(text):
    """order-independent name key, so 'JOHN DOE' matches 'DOE JOHN'"""
    words = re.sub(r"[^a-z ]", " ", str(text).lower()).split()
    return " ".join(sorted(words))


def read_raw(path):
    """Read a sheet with no assumptions about where the header is."""
    try:
        return pd.read_excel(path, header=None, dtype=str)
    except Exception:
        # some .xls exports are really HTML tables
        return pd.read_html(path, header=None)[0].astype(str)


def load_table(path, header_keyword):
    """Find the header row (the one containing header_keyword) and return a clean table."""
    raw = read_raw(path)
    header_row = None
    for i in range(min(30, len(raw))):
        if any(header_keyword in norm(c) for c in raw.iloc[i].tolist()):
            header_row = i
            break
    if header_row is None:
        raise ValueError(f"Could not find a header row containing '{header_keyword}' in {path.name}")

    df = raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[header_row]]
    return df.reset_index(drop=True)


def pick_col(df, keyword, exact=False):
    for col in df.columns:
        n = norm(col)
        if (n == keyword) if exact else (keyword in n):
            return col
    raise KeyError(f"No column matching '{keyword}'. Columns found: {list(df.columns)}")


# ---------- load files ----------
control_file = find_file(["CIVE Control Numbers*.xls*", "CIVE*.xls*"])
class_file = find_file(["DS*YEAR12024*.xlsx", "DS*.xlsx"])
print(f"Control numbers file : {control_file.name}")
print(f"Class list file      : {class_file.name}\n")

control = load_table(control_file, "billitemref")
classlist = load_table(class_file, "regist")

# control numbers file
c_name = pick_col(control, "payername")
c_reg = pick_col(control, "billitemref")

# class list file ("regist" also matches the typo REGISTTRATION)
k_name = pick_col(classlist, "name", exact=True)
k_reg = pick_col(classlist, "regist")
k_prog = pick_col(classlist, "programme")

control = control[[c_name, c_reg]].dropna(how="all")
control["reg_key"] = control[c_reg].map(norm_reg)
control["name_key"] = control[c_name].map(norm_name)
control = control[control["reg_key"].str.len() > 0].drop_duplicates("reg_key")

classlist = classlist[[k_name, k_reg, k_prog]].dropna(how="all").copy()
classlist["reg_key"] = classlist[k_reg].map(norm_reg)
classlist["name_key"] = classlist[k_name].map(norm_name)

# ---------- match ----------
by_reg = classlist["reg_key"].isin(control["reg_key"])
matched = classlist[by_reg]

# fallback: people whose reg number didn't match but whose name does
leftover_control = control[~control["reg_key"].isin(matched["reg_key"])]
by_name = classlist[~by_reg & classlist["name_key"].isin(leftover_control["name_key"])]

result = pd.concat([matched, by_name])
result = result[[k_name, k_reg, k_prog]]
result.columns = ["NAME", "REGISTRATION NUMBER", "PROGRAMME"]
result = result.sort_values(["PROGRAMME", "NAME"]).reset_index(drop=True)
result.index += 1

# ---------- report ----------
print(f"Extended accommodation students found in class list: {len(result)}\n")
print(result.to_string())

# people in the control numbers file who could not be found in the class list
found_keys = set(matched["reg_key"]) | set(
    control[control["name_key"].isin(by_name["name_key"])]["reg_key"]
)
missing = control[~control["reg_key"].isin(found_keys)]
if len(missing):
    print(f"\n⚠ {len(missing)} entries in the control numbers file were NOT found in the class list:")
    print(missing[[c_name, c_reg]].to_string(index=False))

result.to_excel(OUTPUT, index_label="#")
print(f"\nSaved to: {OUTPUT.name}")