# Walkthrough: Porting VBA Logic to React Application

I have successfully ported the core logic from the "TOOL_160925_BOQ_BOM_DNO(3 Rail).xlsm" VBA macros into your React application using Clean Architecture.

## Key Achievements

1.  **Equipment Data Extraction**:
    - Created a PowerShell script (`extract_data_com_fast.ps1`) to securely extract Inverter and PV Module data from the `DATA EQUIP` sheet of your Excel file.
    - Generated `src/data/equipment_data.json` containing detailed specifications for 13 Huawei Inverters and 9 PV Models.

2.  **VBA Logic Translation**:
    - Implemented `SelectHuaweiInverter` function in JavaScript, matching the logic provided (kWp-based selection).
    - Ported the "3-Rail" estimation logic for rails, clamps, splices, and L-feet.
    - Implemented logic for calculating cable lengths, MC4 connectors, and grounding materials based on project parameters.
    - Grouped the output items into the standard 8 groups (I. Structure, II. PV, III. Inverter, etc.) as per your BOQ format.

3.  **UI Updates**:
    - Updated `ProjectConfig` to support `PV Model` and `Inverter Position` selection.
    - Enhanced `CalculatorPage` to:
        - Support manual entry of PV Model and Inverter Position.
        - Display results in a clear, grouped table format.
        - Aggregate costs (placeholder) and quantities.
    - Updated `DatalinkImporter` to automatically populate default values for new fields.

## Files Created/Modified

-   **`extract_data_com_fast.ps1`**: Utility to re-extract data from Excel if needed.
-   **`src/data/equipment_data.json`**: The database of equipment.
-   **`src/domain/usecases/CalculateBOQ.js`**: The core logic engine (Rewrite).
-   **`src/presentation/pages/CalculatorPage.jsx`**: The main interface for BOQ generation.
-   **`src/domain/entities/ProjectConfig.js`**: Updated entity.
-   **`src/presentation/components/DatalinkImporter.jsx`**: updated importer.

## Verification

I validated the logic using a test script which confirmed:
-   Correct Inverter selection (e.g., 50kWp -> `SUN2000-50KTL-M3`).
-   Correct Rail quantity calculation using the formula `(RailCount + 2) / 2` (or adjusted logic).
-   Correct MC4 connector logic based on Inverter inputs.

The application is now ready to generate BOQs that match your Excel tool's output logic.
