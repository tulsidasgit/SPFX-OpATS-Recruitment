import { SPFI } from '@pnp/sp';
import { GraphService } from './shared/GraphService';
import { ICurrentUser } from './shared/models';

export interface IRecruitmentTrackerProps {
  sp: SPFI;
  graphService: GraphService;
  currentUser: ICurrentUser;
  applyBaseUrl: string;
}
