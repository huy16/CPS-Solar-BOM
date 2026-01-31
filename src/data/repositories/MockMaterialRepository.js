import { IMaterialRepository } from '../../domain/interfaces/IMaterialRepository';
import { Material } from '../../domain/entities/Material';

export class MockMaterialRepository extends IMaterialRepository {
    constructor() {
        super();
        this.materials = [
            new Material('1', 'Solar Rail 4.4m', 'Rail', 'bar', 450000),
            new Material('2', 'Mid Clamp', 'Clamp', 'pcs', 12000),
            new Material('3', 'End Clamp', 'Clamp', 'pcs', 12000),
            new Material('4', 'L-Foot', 'Connector', 'pcs', 25000),
            new Material('5', 'M8x25 Bolt', 'Bolt', 'pcs', 2000),
        ];
    }

    async getAllMaterials() {
        // Simulate network delay
        return new Promise(resolve => setTimeout(() => resolve(this.materials), 500));
    }

    async getMaterialById(id) {
        const found = this.materials.find(m => m.id === id);
        return Promise.resolve(found || null);
    }

    async getMaterialsByType(type) {
        return Promise.resolve(this.materials.filter(m => m.type === type));
    }
}
