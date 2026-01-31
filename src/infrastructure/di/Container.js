import { CalculateBOQ } from '../../domain/usecases/CalculateBOQ';
import { RealMaterialRepository } from '../../data/repositories/RealMaterialRepository';

// Singleton instances
const materialRepository = new RealMaterialRepository();
const calculateBOQUseCase = new CalculateBOQ(materialRepository);

export const DI = {
    resolveCalculateBOQ: () => calculateBOQUseCase,
    resolveMaterialRepository: () => materialRepository
};
