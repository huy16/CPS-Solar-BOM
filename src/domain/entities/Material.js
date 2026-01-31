/**
 * Entity representing a physical material or component.
 */
export class Material {
    constructor(id, name, type, unit, unitPrice, specs = {}) {
        this.id = id;
        this.name = name;
        this.type = type; // e.g., 'Rail', 'Clamp', 'Bolt'
        this.unit = unit; // e.g., 'pcs', 'm', 'kg'
        this.unitPrice = unitPrice;
        this.specs = specs; // Additional technical specifications
    }

    /**
     * Calculate cost for a given quantity
     * @param {number} quantity 
     * @returns {number}
     */
    calculateCost(quantity) {
        return this.unitPrice * quantity;
    }
}
