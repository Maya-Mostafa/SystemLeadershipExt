import { ITask} from '../Interfaces/ITask';
import spservice from '../Services/spservices';
export interface INewTaskProps{
  spservice: spservice;
  displayDialog:boolean;
  onDismiss: (refresh:boolean) => void;
}
