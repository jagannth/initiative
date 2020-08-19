import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IDropdownOption } from 'office-ui-fabric-react';
import { ISubmission } from './ISubmission';

export interface IInitiativeProgressSubmissionProps {
  description?: string;
  context?:WebPartContext;
  itemcount?:number;
  userCount?:number;
  Programs?:IDropdownOption[];
  Items?:ISubmission[];
}
