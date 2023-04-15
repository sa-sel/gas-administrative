import { StudentBasicModel } from '@lib';

export const enum ProjectRole {
  Coordinator = 'Coordenação',
  Director = 'Direção',
  Manager = 'Gerência',
  Member = 'Equipe',
}

export type ProjectMemberModel = StudentBasicModel & { role: ProjectRole };
