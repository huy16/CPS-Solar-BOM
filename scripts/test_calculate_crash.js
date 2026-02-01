import { CalculateBOQ } from '../src/domain/usecases/CalculateBOQ.js';
import { ProjectConfig } from '../src/domain/entities/ProjectConfig.js';

// Mock Repository (not used in current logic but required by constructor)
const mockRepo = {};

async function runTest() {
    console.log("Starting Test Calculation...");

    const calculateUseCase = new CalculateBOQ(mockRepo);

    // 1. Valid Test Case
    const validConfig = new ProjectConfig(
        "TEST-001",
        "Test Site A",
        15, // kWp
        22, // Panel Count
        "3-Rail",
        "HSM-ND66-GK715",
        "Lap canh tu MSB"
    );

    try {
        console.log("Testing Valid Config...");
        const result = await calculateUseCase.execute(validConfig);
        console.log("Valid Config Result:", JSON.stringify(result, null, 2));
    } catch (error) {
        console.error("Valid Config FAILED:", error);
    }

    // 2. Edge Case: Zero Power (Should handle gracefully or throw known error)
    const zeroPowerConfig = new ProjectConfig(
        "TEST-002",
        "Zero Power Site",
        0,
        0,
        "3-Rail",
        "HSM-ND66-GK715",
        "Lap canh tu MSB"
    );

    try {
        console.log("\nTesting Zero Power Config...");
        const result = await calculateUseCase.execute(zeroPowerConfig);
        console.log("Zero Power Result:", result);
    } catch (error) {
        console.log("Zero Power Expected Error:", error.message);
    }

    // 3. Edge Case: Unknown PV Model
    const unknownPVConfig = new ProjectConfig(
        "TEST-003",
        "Unknown PV Site",
        10,
        15,
        "3-Rail",
        "UNKNOWN-MODEL",
        "Lap canh tu MSB"
    );

    try {
        console.log("\nTesting Unknown PV Config...");
        const result = await calculateUseCase.execute(unknownPVConfig);
        console.log("Unknown PV Result:", result);
    } catch (error) {
        console.log("Unknown PV Expected Error:", error.message);
    }
}

runTest();
