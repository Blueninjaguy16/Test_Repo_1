import pandas as pd

def flatten_taxonomy(product_box_file, box_part_file, output_file):
    # Step 1: Load Excel sheets, skipping the header row
    df_prod = pd.read_excel(product_box_file, header=None, skiprows=1, usecols="A:J", dtype=str)
    df_box = pd.read_excel(box_part_file, header=None, skiprows=1, dtype=str)  # ← updated line


    # Step 2: Build Box → Parts mapping
    box_to_parts = {}
    for _, row in df_box.iterrows():
        box_id = str(row[0]).strip()
        if not box_id:
            continue
        parts = [str(cell).strip() for cell in row[1:] if pd.notna(cell) and str(cell).strip()]
        box_to_parts[box_id] = parts

    # Step 3: Build Product → Boxes mapping
    product_to_boxes = {}
    for _, row in df_prod.iterrows():
        product_id = str(row[0]).strip()
        if not product_id:
            continue
        boxes = [str(cell).strip() for cell in row[1:] if pd.notna(cell) and str(cell).strip()]
        product_to_boxes[product_id] = boxes

    # Step 4: Determine maximum number of Boxes and Parts
    max_boxes = max((len(b) for b in product_to_boxes.values()), default=0)
    max_parts = max((len(p) for p in box_to_parts.values()), default=0)

    # Step 5: Build flattened hierarchy with aligned columns
    output_rows = []

    for product_id in product_to_boxes:
        boxes = product_to_boxes[product_id]
        row = [product_id, "Product"] + boxes
        output_rows.append(row)

        for box_id in boxes:
            parts = box_to_parts.get(box_id, [])
            row = [box_id, "Box"] + [''] * max_boxes + parts
            output_rows.append(row)

            for part_id in parts:
                output_rows.append([part_id, "Part"])

    # Step 6: Create headers
    header = ["Unique ID", "Taxonomy"]
    header += ["Boxes"] * max_boxes
    header += ["Parts"] * max_parts

    # Step 7: Write DataFrame with correct column count
    df_out = pd.DataFrame(output_rows)
    while df_out.shape[1] < len(header):
        df_out[df_out.shape[1]] = ''
    df_out.columns = header[:df_out.shape[1]]

    # Step 8: Save to Excel
    df_out.to_excel(output_file, index=False)
    print(f"✅ Finished. {len(df_out)} rows written to {output_file}")

# --- Change these to your actual file paths ---
flatten_taxonomy(
    "C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports/product-box.xlsx",
    "C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports/box-part.xlsx",
    "C:/Users/panderson/OneDrive - American Bath Group/Documents/Reports/output.xlsx"
)
