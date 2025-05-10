import { PermissionScope } from '../../enums/permission-scope.enum';

export interface StateObject {
  userId: string;
  csrf: string;
  timestamp: number;
  requestedScopes?: PermissionScope[];
}
