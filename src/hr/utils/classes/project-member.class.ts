import { ProjectRole } from '@hr/models';
import { Student } from '@lib';

export class ProjectMember extends Student {
  role: ProjectRole = ProjectRole.Member;
}
