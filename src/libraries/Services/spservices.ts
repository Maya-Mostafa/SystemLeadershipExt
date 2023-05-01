import * as moment from 'moment';
import { IGroup } from '../Interfaces/IGroups';
import { IGroupMember, IMember } from '../Interfaces/IGroupMembers';
import { IPlannerBucket } from '../Interfaces/IPlannerBucket';
import { IPlannerPlan } from '../Interfaces/IPlannerPlan';
import { IPlannerPlanDetails } from '../Interfaces/IPlannerPlanDetails';
import { IPlannerPlanExtended } from '../Interfaces/IPlannerPlanExtended';
import { ITask } from '../Interfaces/ITask';
import { ITaskDetails } from '../Interfaces/ITaskDetails';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "@pnp/graph/planner";
import {
  sp,
  Web,
  PagedItemCollection,
  ChunkedFileUploadProgressData,
  FileAddResult,
  TaskAddResult
} from '@pnp/pnpjs';

import { spfi, SPFx as spSPFx} from "@pnp/sp";
import { graphfi, SPFx as graphSPFx, DefaultHeaders} from "@pnp/graph";
import "@pnp/graph/planner";
import "@pnp/graph/groups";
import "@pnp/graph/users";
import "@pnp/graph/photos";

// import { SPComponentLoader } from '@microsoft/sp-loader';

const DEFAULT_PERSONA_IMG_HASH: string = '7ad602295f8386b7615b582d87bcc294';
const DEFAULT_IMAGE_PLACEHOLDER_HASH: string = '4a48f26592f4e1498d7a478a4c48609c';
const MD5_MODULE_ID: string = '8494e7d7-6b99-47b2-a741-59873e42f16f';
const PROFILE_IMAGE_URL: string = '/_layouts/15/userphoto.aspx?size=M&accountname=';

export default class spservices {
  
  private graphClient: MSGraphClientV3 = null;
  public currentUser: string = undefined;

  constructor(public context: any, public msGraphClientFactory: any) {
    /*
    Initialize MSGraph
    */

    //const sp = spfi().using(spSPFx(this.context));
    console.log("pageContext", this.context)
    this.currentUser = this.context._user.email;
    this.msGraphClientFactory = this.msGraphClientFactory;
  }

  /**
   * Gets user
   * @param userId
   * @returns user
   */
  public async getUser(userId: string): Promise<IMember> {
    try {
      const graph = graphfi().using(graphSPFx(this.context));
      const user: IMember = await graph.users.getById(userId)();
      return user;
    } catch (error) {
      throw new Error('Error on get user details');
    }
  }

  /**
   * Gets group members
   * @param groupId
   * @returns group members
   */
  public async getGroupMembers(groupId: string, skipToken: string): Promise<IGroupMember> {
    try {
      let groupMembers: IGroupMember;
      if (skipToken) {
        this.graphClient = await this.msGraphClientFactory.getClient('3');
        groupMembers = await this.graphClient
          .api(`groups/${groupId}/members`)
          .version('v1.0')
          .skipToken(skipToken)
          .get();
      } else {
        this.graphClient = await this.msGraphClientFactory.getClient('3');
        groupMembers = await this.graphClient
          .api(`groups/${groupId}/members`)
          .version('v1.0')
          .top(100)
          .get();
      }
      return groupMembers;
    } catch (error) {
      throw new Error('Error on get group members');
    }
  }

  /**
   * Searchs users
   * @param searchString
   * @returns users
   */
  public async searchUsers(searchString: string): Promise<IMember[]> {
    try {
      this.graphClient = await this.msGraphClientFactory.getClient('3');
      const returnUsers = await this.graphClient
        .api(`users`)
        .version('v1.0')
        .top(100)
        .filter(`startswith(DisplayName, '${searchString}') or startswith(mail, '${searchString}')`)
        .get();

      return returnUsers.value;
    } catch (error) {
      throw new Error('Error on search users');
    }
  }

  /**
   * Adds task
   * @param taskInfo
   * @returns task
   */
  public async addTask(taskInfo: string[]): Promise<TaskAddResult> {
    try {
      this.graphClient = await this.msGraphClientFactory.getClient('3');
      const task = await this.graphClient
        .api(`planner/tasks`)
        .version('v1.0')
        .post({
          planId: taskInfo['planId'],
          bucketId: taskInfo['bucketId'],
          title: taskInfo['title'],
          dueDateTime: taskInfo['dueDate'] ? moment(taskInfo['dueDate']).toISOString() : undefined,
          assignments: taskInfo['assignments']
        });

      //const task: TaskAddResult = await graph.planner.tasks.add( taskInfo['planId'], taskInfo['title'], taskInfo['assignments'], taskInfo['bucketId']);
      return task;
    } catch (error) {
      throw new Error('Error on add task');
    }
  }

  /**
   * Gets plan buckets
   * @param planId
   * @returns plan buckets
   */
  public async getPlanBuckets(planId: string): Promise<IPlannerBucket[]> {
    try {
      const graph = graphfi().using(graphSPFx(this.context));
      //const plan = await graph.planner.plans.getById(planId);
      //const plannerBuckets: IPlannerBucket[] = await plan.buckets.get();
      const plannerBuckets: IPlannerBucket[] = await graph.planner.plans.getById(planId).buckets();
      return plannerBuckets;
    } catch (error) {
      throw new Error('Error get Planner buckets');
    }
  }

  /**
   * Gets user groups
   * @returns user groups
   */
  public async getUserGroups(): Promise<IGroup[]> {
    let o365Groups: IGroup[] = [];
    try {
      this.graphClient = await this.msGraphClientFactory.getClient('3');
      const groups = await this.graphClient
        .api(`me/memberof`)
        .version('v1.0')
        .get();
      // Get O365 Groups
      for (const group of groups.value as IGroup[]) {
        const hasO365Group = group.groupTypes && group.groupTypes.length > 0 ? group.groupTypes.indexOf('Unified') : -1;
        if (hasO365Group !== -1) {
          o365Groups.push(group);
        }
      }
      return o365Groups;
    } catch (error) {
      Promise.reject(error);
    }
  }

  /**
   * Gets user plans by group id
   * @param groupId
   * @returns user plans by group id
   */
  public async getUserPlansByGroupId(groupId: string): Promise<IPlannerPlan[]> {
    try {
      // const graph = graphfi().using(graphSPFx(this.context));
      const graph = graphfi().using(DefaultHeaders());
      const groupPlans: IPlannerPlan[] = await graph.groups.getById(groupId)();
      console.log("groupPlans", groupPlans);
      return groupPlans;
    } catch (error) {
      Promise.reject(error);
    }
  }

  /**
   * Gets user plans
   * @returns user plans
   */
  public async getUserPlans(): Promise<IPlannerPlanExtended[]> {
    try {
      let userPlans: IPlannerPlanExtended[] = [];
      const o365Groups: IGroup[] = await this.getUserGroups();
      for (const group of o365Groups) {
        const plans: IPlannerPlan[] = await this.getUserPlansByGroupId(group.id);
        for (const plan of plans) {
          const groupPhoto: string = await this.getGroupPhoto(group.id);
          const userPlan: IPlannerPlanExtended = { ...plan, planPhoto: groupPhoto };
          userPlans.push(userPlan);
        }
      }
      // Sort plans by Title
      userPlans = userPlans.sort((a, b) => {
        if (a.title < b.title) return -1;
        if (a.title > b.title) return 1;
        return 0;
      });
      console.log("userPlans", userPlans);
      return userPlans;
    } catch (error) {
      Promise.reject(error);
    }
  }

  /**
   * Gets group photo
   * @param groupId
   * @returns group photo
   */
  public async getGroupPhoto(groupId: string): Promise<any> {
    return new Promise(async (resolve, reject) => {
      try {
        let url: any = '';
        const graph = graphfi().using(graphSPFx(this.context));
        const photo = await graph.groups.getById(groupId).photo.getBlob();
        let reader = new FileReader();

        reader.addEventListener(
          'load',
          () => {
            url = reader.result; // data url
            resolve(url);
          },
          false
        );
        reader.readAsDataURL(photo); // converts the blob to base64 and calls onload
      } catch (error) {
        resolve(undefined);
      }
    });
  }

}
