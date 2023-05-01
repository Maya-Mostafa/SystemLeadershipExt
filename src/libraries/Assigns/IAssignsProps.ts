import { IPlannerPlan } from '../Interfaces/IPlannerPlan';
import { ITask } from '../Interfaces/ITask';
import spservices from '../Services/spservices';
import { IPlannerPlanExtended } from '../Interfaces/IPlannerPlanExtended';
import { AssignMode} from './EAssignMode';
import { IMember } from '../Interfaces/IGroupMembers';

export interface IAssignsProps {

  onDismiss: (assigns?:IMember[]) => void;
  target?: HTMLElement;
  task?: ITask;
  plannerPlan:IPlannerPlanExtended;
  spservice: spservices;
  AssignMode?:  AssignMode;
  assigns?: IMember[];
}
