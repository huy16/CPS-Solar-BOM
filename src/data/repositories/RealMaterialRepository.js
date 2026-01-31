import { IMaterialRepository } from '../../domain/interfaces/IMaterialRepository';
import materialsData from '../sources/materials.json'; // Importing the JSON directly

export class RealMaterialRepository extends IMaterialRepository {
    constructor() {
        super();
        this.materials = materialsData;
    }

    async getAllMaterials() {
        return Promise.resolve(this.materials);
    }

    async getMaterialById(id) {
        const found = this.materials.find(m => m.id === id);
        return Promise.resolve(found || null);
    }

    async getMaterialsByType(type) {
        return Promise.resolve(this.materials.filter(m => m.type === type));
    }

    async getMaterialByName(nameSubstring) {
        return Promise.resolve(this.materials.find(m => m.name.toLowerCase().includes(nameSubstring.toLowerCase())));
    }
}
