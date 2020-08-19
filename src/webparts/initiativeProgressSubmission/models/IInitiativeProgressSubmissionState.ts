import { ISubmission } from './ISubmission';
import { IDropdownOption, IColumn, IGroup } from "office-ui-fabric-react";
import { Counts } from './Counts';

export interface IInitiativeProgressSubmissionState {
  Programs:IDropdownOption[];
  Initiatives: IDropdownOption[];
  Choices: IDropdownOption[];
  errors: {};
  project: ISubmission;
  status: string;
  showform: boolean;
  showEditform: boolean;
  items: ISubmission[];
  selectionDetails: string;
  selectedcount: number;
  userCount: number;
  allCount: number;
  elementId: string;
  userId: number;
  isModalSelection: boolean;
  announcedMessage?: string;
  columns: IColumn[];
  selectedItems: ISubmission;
  groups:IGroup[];
  groupLabels: string[];
  autherId: number;
  dashboardCounts: Counts;
}
