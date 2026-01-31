/**
 * Interface for Material Data Access
 * (In JS we don't have true interfaces, using a class to document the contract)
 */
export class IMaterialRepository {
    async getAllMaterials() {
        throw new Error("Method not implemented");
    }

    async getMaterialById(id) {
        throw new Error("Method not implemented");
    }

    async getMaterialsByType(type) {
        throw new Error("Method not implemented");
    }
}
