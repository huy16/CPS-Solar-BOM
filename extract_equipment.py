import pandas as pd
import json
import os

EXCEL_PATH = r"D:\TOOL GOOGLE ANTIGRAVITY\6. Tool BOQ & BOM\1. Database\TOOL_160925_BOQ_BOM_DNO(3 Rail).xlsm"
OUTPUT_PATH = r"D:\TOOL GOOGLE ANTIGRAVITY\6. Tool BOQ & BOM\boq-bom-app\src\data\equipment_data.json"

def extract_data():
    print(f"Reading {EXCEL_PATH}...")
    try:
        # Load specific sheet to avoid reading entire file if possible, but pandas reads all usually?
        # engine='openpyxl' supports xlsm
        df = pd.read_excel(EXCEL_PATH, sheet_name='DATA EQUIP', engine='openpyxl')
        
        print("Columns:", df.columns.tolist())
        
        equipment_data = {
            "inverters": {},
            "photovoltaics": {}
        }
        
        # Fuzzy column matching helper
        def get_col(name):
            for col in df.columns:
                if str(col).strip().lower() == name.lower():
                    return col
            return None

        # Inverter Columns
        col_inv_model = get_col("Inv Model")
        col_inv_min = get_col("Min Voltage Operating")
        col_inv_max = get_col("Max voltage Operating")
        col_inv_inputs = get_col("Num Inputs")
        col_inv_cap = get_col("Capacity")

        # PV Columns
        col_pv_model = get_col("PV Model")
        # VBA said 'Capacity' (AM) and 'Voc' (AP).
        # We can try to find them by name if they have headers, or by index if headers are missing/duplicate
        # Assuming headers exist since valid generic columns were found.
        # If headers are like "Capacity.1" (duplicate), we need care.
        
        # Let's iterate rows
        for index, row in df.iterrows():
            # Inverter
            if col_inv_model and pd.notna(row[col_inv_model]):
                model = str(row[col_inv_model]).strip()
                if model:
                    equipment_data["inverters"][model] = {
                        "minVoltage": row[col_inv_min] if col_inv_min else 0,
                        "maxVoltage": row[col_inv_max] if col_inv_max else 0,
                        "numInputs": row[col_inv_inputs] if col_inv_inputs else 0,
                        "capacityKw": row[col_inv_cap] if col_inv_cap else 0
                    }
            
            # PV
            if col_pv_model and pd.notna(row[col_pv_model]):
                model = str(row[col_pv_model]).strip()
                if model:
                    # Need to find Capacity and Voc columns.
                    # If pandas headers are unique, maybe "Capacity" is the inverter one. 
                    # PV attributes might be unnamed or duplicate.
                    # Let's rely on finding "Voc" at least.
                    
                    # Hardcoded indices fallback if names fail?
                    # Pandas removes empty top rows usually.
                    # Column AM is index 38. AP is index 41.
                    
                    try:
                        cap = row.iloc[38] # Index 38 (Column AM)
                        voc = row.iloc[41] # Index 41 (Column AP)
                        
                        # Handle potential NaN
                        cap = float(cap) if pd.notna(cap) else 0
                        voc = float(voc) if pd.notna(voc) else 0
                        
                        equipment_data["photovoltaics"][model] = {
                            "powerKwp": cap / 1000,
                            "voc": voc
                        }
                    except Exception as e:
                        # print(f"Error extracting PV data for {model}: {e}")
                        pass

        # Save
        os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
        with open(OUTPUT_PATH, 'w') as f:
            json.dump(equipment_data, f, indent=2)
            
        print(f"Saved to {OUTPUT_PATH}")
        print(f"Inverters: {len(equipment_data['inverters'])}")
        print(f"PVs: {len(equipment_data['photovoltaics'])}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    extract_data()
