import pandas as pd

# === Load Excel files ===
backup_df = pd.read_excel(
    r"C:\Users\panderson\OneDrive - American Bath Group\Documents\Paul_Anderson\Project_Archive\DC\DC0234\DC0234-backup-5.9.25.xlsx"
)
import_df = pd.read_excel(
    r"C:\Users\panderson\OneDrive - American Bath Group\Documents\Paul_Anderson\Project_Archive\DC\DC0234\DC0234-import-5.9.25.xlsx"
)

# === Normalize string data ===
def normalize(df):
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].fillna('').astype(str).str.strip().str.lower()
    return df

# === Inherit 'Brand' from last Base row ===
def inherit_base_field(df, id_col, fields_to_inherit):
    last_base = {}
    rows = []
    for _, row in df.iterrows():
        if row[id_col] == 'base':
            for f in fields_to_inherit:
                last_base[f] = row[f]
        new = row.copy()
        for f in fields_to_inherit:
            if new[f] == '':
                new[f] = last_base.get(f, '')
        rows.append(new)
    return pd.DataFrame(rows)

# === Apply normalization & inheritance ===
backup_df = normalize(backup_df)
import_df = normalize(import_df)

fields_to_inherit = ['Brand']
hierarchy_col = 'salsify:data_inheritance_hierarchy_level_id'
backup_df = inherit_base_field(backup_df, hierarchy_col, fields_to_inherit)
import_df = inherit_base_field(import_df, hierarchy_col, fields_to_inherit)

# === Identify Unique ID column ===
def find_column(df, name_match):
    return next(col for col in df.columns if col.strip().lower() == name_match)

unique_id_col = find_column(backup_df, 'unique id')

# === Define your priority columns and weights ===
# Matches on these count heavier than others
priority_cols = ['Base Part Number', 'Sellable Part Number', 'Brand','Color']
priority_weight = 10      # each match here adds +10
# all other shared columns get +1
compare_cols = [
    col for col in import_df.columns
    if col in backup_df.columns and col.strip().lower() != 'unique id'
]

# === Prepare output ===
import_df['Matched Unique ID'] = ''

# === Matching loop with weights ===
for i, imp in import_df.iterrows():
    best_score = -1
    best_uid = ''
    for _, bkp in backup_df.iterrows():
        score = 0
        # weighted matches
        for col in priority_cols:
            if col in compare_cols and imp[col] == bkp[col]:
                score += priority_weight
        # normal matches
        for col in compare_cols:
            if col not in priority_cols and imp[col] == bkp[col]:
                score += 1
        if score > best_score:
            best_score = score
            best_uid = bkp[unique_id_col]
    import_df.at[i, 'Matched Unique ID'] = best_uid

# === Save result ===
import_df.to_excel("import_with_matched_ids.xlsx", index=False)
