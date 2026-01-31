import { CalculateBOQ } from '../src/domain/usecases/CalculateBOQ.js';
import { ProjectConfig } from '../src/domain/entities/ProjectConfig.js';

// Simulate the logic from BulkInputComponent
async function parseAndCalculate() {
    console.log("--- Simulating Input Parsing ---");
    const rawInput = "14100\tBHX_DNA_TKH - 198 Nguyễn Phước Nguyên\tCAS-BHX_DNA_14100\t14.16 kWp";
    const cols = rawInput.split('\t');

    console.log("Raw Columns:", cols);

    let id, name, power;

    if (cols.length === 4 && !isNaN(parseFloat(cols[3]))) {
        id = cols[0];
        name = cols[1];
        power = parseFloat(cols[3]);
    } else {
        console.error("Parsing failed!");
        return;
    }

    console.log(`Parsed -> ID: ${id}, Name: ${name}, Power: ${power} kWp`);

    // Defaults
    const config = new ProjectConfig(
        id,
        name,
        power,
        0, // panelCount (calculated from power)
        "3-Rail", // default assumption for now if not specified? Or standard default?
        "HSM-ND66-GK715",
        "Lap canh tu MSB",
        10,
        100,
        50
    );

    console.log("--- Running Calculation ---");
    const calculator = new CalculateBOQ();
    const result = await calculator.execute(config);

    console.log(`Calculated ${result.items.length} items.`);

    // Group and show summary
    const summary = {};
    result.items.forEach(item => {
        if (!summary[item.name]) summary[item.name] = 0;
        summary[item.name] += item.quantity;
    });

    console.log("\n--- BOQ Result Summary ---");
    console.table(result.items.map(i => ({ Group: i.group, Code: i.code, Name: i.name, Qty: i.quantity, Unit: i.unit })));
}

parseAndCalculate();
