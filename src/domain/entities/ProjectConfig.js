/**
 * Entity representing the Project Configuration inputs.
 */
export class ProjectConfig {
    constructor(id, name, dcPower, panelCount, railSystemType, pvModel, inverterPosition, acCable, dcCable, cat5Cable) {
        this.id = id || Date.now().toString();
        this.name = name;
        this.dcPower = dcPower; // in kWp
        this.panelCount = panelCount;
        this.railSystemType = railSystemType; // e.g., '3-Rail', '2-Rail'
        this.pvModel = pvModel || "HSM-ND66-GK715"; // Default from VBA
        this.inverterPosition = inverterPosition || "Lap canh tu MSB"; // Default from VBA
        this.acCable = acCable || 10;
        this.dcCable = dcCable || 80;
        this.cat5Cable = cat5Cable || 40;
    }

    isValid() {
        return this.dcPower > 0;
    }
}
