import * as React from 'react';
import { useId, useBoolean } from '@uifabric/react-hooks';
//import styles from './InitiativeProgressSubmission.module.scss';
import { IInitiativeProgressSubmissionProps } from '../models/IInitiativeProgressSubmissionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './styles.module.scss';
import { IInitiativeProgressSubmissionState } from '../models/IInitiativeProgressSubmissionState';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import { IStyleSet } from 'office-ui-fabric-react/lib/Styling';

import { MySubmissions } from '../components/MySubmissions';
import { Dashboard } from '../components/Dashboard';
import { Reports } from '../components/Reports';
import { NeedAttention } from '../components/NeedAttention';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ISubmission } from '../models/ISubmission';
import { ModalBasicExample } from './AddNewModal';
import { Crudoperations } from '../services/SPServices';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import {
  IDropdownOption,
  Fabric,
  CommandButton, Panel,
  PanelType, PrimaryButton,
  DefaultButton, TextField,
  Dropdown, Announced,
  MarqueeSelection,
  IRenderFunction
} from 'office-ui-fabric-react';
import {
  DetailsList,
  DetailsHeader,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  IGroup,
  IDetailsHeaderProps
} from 'office-ui-fabric-react/lib/DetailsList';

import {
  getTheme,
  mergeStyleSets,
  FontWeights,
  ContextualMenu,
  Toggle,
  Modal,
  IDragOptions,
  IconButton,
  IIconProps,
} from 'office-ui-fabric-react';
import { Counts } from '../models/Counts';



const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
  root: { marginTop: 10 },
};
export default class InitiativeProgressSubmission extends React.Component<IInitiativeProgressSubmissionProps, IInitiativeProgressSubmissionState, {}> {
  public _spops: Crudoperations;
  private _selection: Selection;



  constructor(props: IInitiativeProgressSubmissionProps) {
    super(props);
    SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css");
    SPComponentLoader.loadScript("https://kit.fontawesome.com/74a9a9044f.js");
    this._spops = new Crudoperations();
    //this.props.context; 
    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'Initiative',
        fieldName: 'Initiative',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.intiate}</span>;
        },
      },
      {
        key: 'column2',
        name: 'Programs',
        fieldName: 'Programs',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.program}</span>;
        },
      },
      {
        key: 'column3',
        name: 'key achievements in period',
        fieldName: 'keyachievementsinperiod',
        minWidth: 210,
        maxWidth: 350,
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.achievements}</span>;
        },
      },
      {
        key: 'column4',
        name: 'key activities for next period',
        fieldName: 'keyactivitiesfornextperiod',
        minWidth: 210,
        maxWidth: 350,
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.activities}</span>;
        },
      },
      {
        key: 'column5',
        name: 'support /attention needed',
        fieldName: 'supportattentionneeded',
        minWidth: 210,
        maxWidth: 350,
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.supportnd}</span>;
        },
      },
      {
        key: 'column6',
        name: 'Scope Status',
        fieldName: 'ScopeStatus',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.scopsts}</span>;
        },
      },
      {
        key: 'column7',
        name: 'Schedule Status',
        fieldName: 'ScheduleStatus',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.schdlsts}</span>;
        },
      },
      {
        key: 'column8',
        name: 'Business Case Status',
        fieldName: 'BusinessCaseStatus',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.bssts}</span>;
        },
      },
      {
        key: 'column9',
        name: 'Overall Status',
        fieldName: 'OverallStatus',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.ovrlsts}</span>;
        },
      },
      {
        key: 'column10',
        name: 'Scope Trend',
        fieldName: 'ScopeTrend',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.scoptrnd}</span>;
        },
      },
      {
        key: 'column11',
        name: 'Schedule Trend',
        fieldName: 'ScheduleTrend',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.schdltrnd}</span>;
        },
      },
      {
        key: 'column12',
        name: 'Business Case Trend',
        fieldName: 'BusinessCaseTrend',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.bstrnd}</span>;
        },
      },
      {
        key: 'column13',
        name: 'Overal Trend',
        fieldName: 'OveralTrend',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.ovrltrnd}</span>;
        },
      },
      {
        key: 'column14',
        name: 'Change & comms Status',
        fieldName: 'ChangecommasStatus',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.cncsts}</span>;
        },
      },
      {
        key: 'column15',
        name: 'Change & comms Trend',
        fieldName: 'ChangecommasTrend',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.cnctrnd}</span>;
        },
      },
      {
        key: 'column16',
        name: 'Impact on Operations Status',
        fieldName: 'ImpactonOperationsStatus',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.imopsts}</span>;
        },
      },
      {
        key: 'column17',
        name: 'Impact on Operations Trend',
        fieldName: 'ImpactonOperationsTrend',
        minWidth: 210,
        maxWidth: 350,
        data: 'string',
        isPadded: true,
        onRender: (item: ISubmission) => {
          return <span>{item.imoptrnd}</span>;
        },
      },
    ];
    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails(),
          selectedItems: this._getselectedItem(),
          selectedcount: this._getitemcount(),
        });
      },
    });
    this.state = {
      Programs: [],
      Initiatives: [],
      Choices: [],
      errors: {},
      project: {},
      status: "",
      showform: false,
      showEditform: false,
      items: [],
      selectionDetails: this._getSelectionDetails(),
      selectedcount: this._getitemcount(),
      userCount: 0,
      allCount: 0,
      elementId: '',
      userId: null,
      isModalSelection: true,
      announcedMessage: undefined,
      columns: columns,
      groups: [],
      groupLabels: [],
      selectedItems: this._getselectedItem(),
      autherId: null,
      dashboardCounts: new Counts()

    };
  }

  public getinitiativeval = (event: any, result: any) => {
    //this.initiate=data.text;
    let x = this.state.project;
    x.intiate = result.text;
    this.setState({
      project: x
    });
  }

  public getprogramval = (event: any, result: any) => {
    //this.initiate=data.text;
    let x = this.state.project;
    x.program = result.text;
    this.setState({
      project: x
    });
    this._spops.getinitiative(this.props.context, this.state.project.program)
      .then((result: IDropdownOption[]) => {
        this.setState({ Initiatives: result });
      });
  }

  public getChoiceOneval = (event: any, result: any) => {
    let x = this.state.project;
    x.Choiceone = result.text;
    this.setState({
      project: x
    });
  }

  public _handlebtnchange = (e, option) => {
    let x = this.state.project;
    let name = e.target.name;
    if (name == "scopeStatus") { x.scopsts = option.text; }
    else if (name == "scheduleStatus") { x.schdlsts = option.text; }
    else if (name == "businessCaseStatus") { x.bssts = option.text; }
    else if (name == "overallStatus") { x.ovrlsts = option.text; }
    else if (name == "scopeTrend") { x.scoptrnd = option.text; }
    else if (name == "scheduleTrend") { x.schdltrnd = option.text; }
    else if (name == "busineeCaseTrend") { x.bstrnd = option.text; }
    else if (name == "overalTrend") { x.ovrltrnd = option.text; }
    else if (name == "cncStatus") { x.cncsts = option.text; }
    else if (name == "cncTrend") { x.cnctrnd = option.text; }
    else if (name == "imOpStatus") { x.imopsts = option.text; }
    else if (name == "imOpTrend") { x.imoptrnd = option.text; }
    this.setState({ project: x });
  }
  public _inputChanged = (e) => {
    let x = this.state.project;
    let txtarea = e.target.name;
    if (txtarea == "achievements") { x.achievements = e.target.value; }
    else if (txtarea == "activities") { x.activities = e.target.value; }
    else if (txtarea == "supportnd") {
      x.supportnd = e.target.value;
    }
    this.setState({
      project: x
    });


  }

  private _onRenderDetailsHeader(props: IDetailsHeaderProps, _defaultRender?: IRenderFunction<IDetailsHeaderProps>) {
    return <DetailsHeader {...props} ariaLabelForToggleAllGroupsButton={'Expand collapse groups'} />;
  }

  private _getitemcount(): number {
    const selectionCount = this._selection.getSelectedCount();
    return selectionCount;
  }

  private _getselectedItem(): ISubmission {
    const selecteditem: ISubmission = {};
    this._selection.getSelection().map((selectitem: any) => {
      //selecteditem.push({
      selecteditem.program = selectitem.program,
        selecteditem.intiate = selectitem.intiate;
      selecteditem.scopsts = selectitem.scopsts;
      selecteditem.schdlsts = selectitem.schdlsts;
      selecteditem.bssts = selectitem.bssts;
      selecteditem.ovrlsts = selectitem.ovrlsts;
      selecteditem.scoptrnd = selectitem.scoptrnd;
      selecteditem.schdltrnd = selectitem.schdltrnd;
      selecteditem.bstrnd = selectitem.bstrnd;
      selecteditem.ovrltrnd = selectitem.ovrltrnd;
      selecteditem.cncsts = selectitem.cncsts;
      selecteditem.cnctrnd = selectitem.cnctrnd;
      selecteditem.imopsts = selectitem.imopsts;
      selecteditem.imoptrnd = selectitem.imoptrnd;
      selecteditem.achievements = selectitem.achievements;
      selecteditem.activities = selectitem.activities;
      selecteditem.supportnd = selectitem.supportnd;
      selecteditem.id = selectitem.id;
      selecteditem.auther = selecteditem.auther;
      //});
    });
    return selecteditem;
  }
  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();
    const selecteditem = this._selection.getSelection();
    //console.log(selecteditem);
    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 selected';
      default:
        return `${selectionCount} selected`;
    }
  }
  private _onRenderColumn(item: ISubmission, index: number, column: IColumn) {
    const value =
      item && column && column.fieldName ? item[column.fieldName as keyof ISubmission] || '' : '';

    return <div data-is-focusable={true}>{value}</div>;
  }

  public componentDidMount() {
    this._spops.getprograms(this.props.context, 'Programs')
      .then((result: IDropdownOption[]) => {
        this.setState({ Programs: result });
      });
    this._spops.getChoicesone(this.props.context)
      .then((result: IDropdownOption[]) => {
        this.setState({ Choices: result });
      });
    this._spops.getlistitems(this.props.context)
      .then((itemResult: ISubmission[]) => {
        console.log('items --', itemResult);
        this.setState({ items: itemResult });
        setTimeout(() => {
          sessionStorage.setItem('initiatives', JSON.stringify(itemResult));
          this._spops.getCurrentUser(this.props.context).then((result: number) => {
            console.log('user', result);
            const usercount = itemResult.filter(x => x.auther == result);
            sessionStorage.setItem('userInitiatives', JSON.stringify(usercount));
            this._dashboardCounts();
            console.log('user count', usercount);
            this.setState({
              userCount: usercount.length,
              allCount: itemResult.length,
              userId: result
            })
          });
        }, 2000);
      });

    this._spops.getGroupingLable(this.props.context)
      .then((result: string[],) => {
        this.setState({ groupLabels: result });
        setTimeout(() => {
          this._spops.getGrouping(result, this.state.items)
            .then((result: IGroup[],) => {
              //console.log('groups --',result);
              this.setState({ groups: result });
            });
        }, 2000);
      });
  }
  public _dashboardCounts() {
    const dashboard = new Counts();
    setTimeout(() => {
      console.log('itemss dashboard', dashboard);
      const itemss = JSON.parse(sessionStorage.getItem('initiatives'));
      const userItemss = JSON.parse(sessionStorage.getItem('userInitiatives'));
      dashboard.All.OnSchedule = itemss.filter(x => x.scopsts == 'On schedule;goals within reach').length;
      dashboard.All.MinurIssues = itemss.filter(x => x.scopsts == 'minor issues threatening scheduleand/or goals').length;
      dashboard.All.NeedHelp = itemss.filter(x => x.scopsts == 'Behind schedule and/or goals are risk').length;
      dashboard.All.TrendingUp = itemss.filter(x => x.scoptrnd == 'Trending up').length;
      dashboard.All.Stable = itemss.filter(x => x.scoptrnd == 'Stable').length;
      dashboard.All.TreningDown = itemss.filter(x => x.scoptrnd == 'Trending down').length;

      dashboard.User.OnSchedule = userItemss.filter(x => x.scopsts == 'On schedule;goals within reach').length;
      dashboard.User.MinurIssues = userItemss.filter(x => x.scopsts == 'minor issues threatening scheduleand/or goals').length;
      dashboard.User.NeedHelp = userItemss.filter(x => x.scopsts == 'Behind schedule and/or goals are risk').length;
      dashboard.User.TrendingUp = userItemss.filter(x => x.scoptrnd == 'Trending up').length;
      dashboard.User.Stable = userItemss.filter(x => x.scoptrnd == 'Stable').length;
      dashboard.User.TreningDown = userItemss.filter(x => x.scoptrnd == 'Trending down').length;
      this.setState({
        dashboardCounts: dashboard
      })
    }, 1000);
  }
  public _initiativeTopFilter(value,id) {
    console.log('working');
    this.setState({ elementId: id});
    console.log('value', value, this.state.userId);
    const itemss = JSON.parse(sessionStorage.getItem('initiatives'));
    const userItemss = JSON.parse(sessionStorage.getItem('userInitiatives'));
    if (value == 'All') {
      this.setState({
        items: itemss
      })
    } else {
      this.setState({
        items: userItemss
      })
    }
    setTimeout(() => {
      console.log('initiatives fromn sess', this.state.items);
      this._spops.getGrouping(this.state.groupLabels, this.state.items)
        .then((result: IGroup[],) => {
          //console.log('groups --',result);
          this.setState({ groups: result });
        });
    }, 1000);
  }
  public _initiativeScopeFilter(value, key,id) {
    this.setState({ showform: false,elementId:id });
    const itemss = JSON.parse(sessionStorage.getItem('initiatives'));
    const userItemss = JSON.parse(sessionStorage.getItem('userInitiatives'));
    if (value == 'All') {
      this.setState({
        items: itemss.filter(x => x.scopsts == key)
      })
    } else {
      this.setState({
        items: userItemss.filter(x => x.scopsts == key)
      })
    }
    setTimeout(() => {
      this._spops.getGrouping(this.state.groupLabels, this.state.items)
        .then((result: IGroup[],) => {
          this.setState({ groups: result });
        });
    }, 1000);
  }
  public _initiativeTrendFilter(value, key,id) {
    this.setState({ showform: false, elementId:id });
    console.log('value', value, this.state.userId);
    const itemss = JSON.parse(sessionStorage.getItem('initiatives'));
    const userItemss = JSON.parse(sessionStorage.getItem('userInitiatives'));
    if (value == 'All') {
      this.setState({
        items: itemss.filter(x => x.scoptrnd == key)
      })
    } else {
      this.setState({
        items: userItemss.filter(x => x.scoptrnd == key)
      })
    }
    setTimeout(() => {
      this._spops.getGrouping(this.state.groupLabels, this.state.items)
        .then((result: IGroup[],) => {
          this.setState({ groups: result });
        });
    }, 1000);
  }
  public render(): React.ReactElement<IInitiativeProgressSubmissionProps> {
    const dragOptions: IDragOptions = {
      moveMenuItemText: 'Move',
      closeMenuItemText: 'Close',
      menu: ContextualMenu,
    };
    const cancelIcon: IIconProps = { iconName: 'Cancel' };
    //const titleId = useId('title');
    const titleId = "Add a new item";
    let option: IDropdownOption[] = [];
    let rbtOption: IChoiceGroupOption;
    const {
      showEditform,
      showform,
      columns,
      items,
      groups,
      selectionDetails,
      isModalSelection,
      announcedMessage,
      selectedItems,
      selectedcount
    } = this.state;
    let x = this.state.project;
    return (
      <div className={styles.initiativeProgressSubmission} style={{ backgroundColor: '#f3f2f1' }}>
        <div className={styles.contianer}>
          <Pivot aria-label="Basic Pivot Example">
            <PivotItem itemIcon="AddWork"
              headerText="Submit Progress Report"
              headerButtonProps={{
                'data-order': 1,
                'data-title': 'Submit Progress Report',
              }}
            >
              <Label styles={labelStyles} className="my-2">
                <div className={styles.padd}>
                  <Fabric>
                    <div className={styles.Header}>

                      {selectedcount ? (
                        <div>
                          <CommandButton
                            onClick={() => { this.setState({ showEditform: !this.state.showEditform }); }}
                            iconProps={{ iconName: 'Edit' }} text="Edit">
                          </CommandButton>
                          <CommandButton
                            onClick={() => this._spops.deleteItem(this.props.context, this._selection.getSelection())
                              .then((results: ISubmission[]) => {
                                this.setState({ items: results });
                                //console.log('group label', this.state.groupLabels);
                                setTimeout(() => {
                                  //console.log('item group', results);
                                  this._spops.getGrouping(this.state.groupLabels, results)
                                    .then((result: IGroup[]) => {
                                      //console.log('resultttttt group', result);
                                      this.setState({ groups: result });
                                    })
                                }, 2000);
                              })}
                            iconProps={{ iconName: 'Delete' }}
                            text="Delete">
                          </CommandButton>
                        </div>
                      ) : (
                          <div>
                            <PrimaryButton className={styles.BtneOne} text="Submit Progress">
                            </PrimaryButton>
                            {/* onClick={() => { this.setState({ showform: true }); }} */}
                          </div>
                        )}
                      <Modal
                        titleAriaId={this.state.selectedItems.intiate}
                        isOpen={this.state.showEditform}
                        isBlocking={false}
                        containerClassName={styles.modal}>
                        <div className={styles.modal} style={{ padding: '20px' }}>
                          <div className={styles.modalinner}>
                            <div className={styles.ModalHeader}>
                              <h5>{this.state.selectedItems.intiate}</h5>
                              <p style={{ color: '#fff' }}>This is the popup about adding Programs,Initiative, Key Achievements in Period,Key Activities for next Period,Support / Attention Needed,Scope , Schedule  etc.</p>
                            </div>
                            <div className={styles.ModalBody}>
                              <div className="dflex mb2">
                                <div className={`${styles.formgroup} ${styles.mr4}`}>
                                  <Label ><strong>Programs :</strong></Label>
                                  <Dropdown
                                    placeholder={selectedItems.program}
                                    options={this.state.Programs}
                                    onChange={this.getprogramval}>
                                  </Dropdown>
                                </div>
                                <div className={styles.formgroup}>
                                  <Label ><strong>Initiative :</strong></Label>
                                  <Dropdown
                                    placeholder={selectedItems.intiate}
                                    options={this.state.Initiatives}
                                    onChange={this.getinitiativeval}>
                                  </Dropdown>
                                </div>
                              </div>
                              <div className={styles.formgroup}>
                                <Label><strong>Key Achievements in Period : </strong></Label>
                                <TextField className={`{$styles.formcontrol} {$styles.bgone}`}
                                  onChange={this._inputChanged.bind(this)} defaultValue={selectedItems.achievements}
                                  name="achievements" required={true} cols={50} rows={5}
                                  multiline
                                  contentEditable={true}
                                ></TextField>
                              </div>

                              <div className={styles.formgroup}>
                                <Label><strong>Key Activities for next Period :</strong></Label>
                                <TextField className={`{$styles.formcontrol} {$styles.bgone}`}
                                  onChange={this._inputChanged.bind(this)}
                                  name="activities" required={true} cols={50} rows={5}
                                  defaultValue={selectedItems.activities} contentEditable={true} multiline ></TextField>
                              </div>

                              <div className={styles.formgroup}>
                                <Label><strong>Support / Attention Needed : </strong></Label>
                                <TextField className={`{$styles.formcontrol} {$styles.bgone}`}
                                  onChange={this._inputChanged.bind(this)}
                                  defaultValue={selectedItems.supportnd}
                                  contentEditable={true}
                                  name="supportnd" required={true} cols={50} rows={5} multiline ></TextField>
                              </div>
                              {/* scope container */}
                              <div className="row">
                                <div className="col-6">
                                  <Label><strong>Scope Status :</strong></Label>
                                  <ChoiceGroup defaultSelectedKey={selectedItems.scopsts}
                                    name="scopeStatus"
                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">
                                  <Label><strong>Scope Trend : </strong></Label>
                                  <ChoiceGroup
                                    name="scopeTrend" defaultSelectedKey={selectedItems.scoptrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                              </div>
                              {/* schedule container */}
                              <div className="row">
                                <div className="col-6">
                                  <Label><strong>Schedule Status : </strong></Label>
                                  <ChoiceGroup defaultSelectedKey={selectedItems.schdlsts}
                                    name="scheduleStatus"
                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">
                                  <Label><strong>Schedule Trend : </strong></Label>
                                  <ChoiceGroup
                                    name="scheduleTrend" defaultSelectedKey={selectedItems.schdltrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                              </div>
                              {/* business container */}
                              <div className="row">
                                <div className="col-6">
                                  <Label><strong>Business Case Status :</strong></Label>
                                  <ChoiceGroup
                                    name="businessCaseStatus" defaultSelectedKey={selectedItems.bssts}

                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">
                                  <Label><strong>Business case Trend :</strong></Label>
                                  <ChoiceGroup
                                    name="businessCaseTrend" defaultSelectedKey={selectedItems.bstrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                              </div>
                              {/* overall container */}
                              <div className="row">
                                <div className="col-6">
                                  <Label><strong>Overall Status : </strong></Label>
                                  <ChoiceGroup
                                    name="overallStatus" defaultSelectedKey={selectedItems.ovrlsts}
                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">

                                  <Label><strong>Overall Trend : </strong></Label>
                                  <ChoiceGroup
                                    name="overallTrend" defaultSelectedKey={selectedItems.ovrltrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                              </div>
                              {/* commission container */}
                              <div className="row">
                                <div className="col-6">

                                  <Label><strong>Communication & Comms. Status : </strong></Label>
                                  <ChoiceGroup
                                    name="cncStatus" defaultSelectedKey={selectedItems.cncsts}
                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">
                                  <Label><strong>Communication & Comms. Trend : </strong></Label>
                                  <ChoiceGroup
                                    name="cncTrend" defaultSelectedKey={selectedItems.cnctrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />

                                </div>
                              </div>
                              {/* impact container */}
                              <div className="row">
                                <div className="col-6">
                                  <Label><strong>Impact on opt Status :</strong></Label>
                                  <ChoiceGroup
                                    name="imOpStatus" defaultSelectedKey={selectedItems.imopsts}
                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">
                                  <Label><strong>Impact on opt Trend : </strong></Label>
                                  <ChoiceGroup
                                    name="imOpTrend" defaultSelectedKey={selectedItems.imoptrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                              </div>
                              <div className={styles.ModalFooter}>
                                <PrimaryButton text="Save" type="submit" className={styles.BtneOne}
                                  onClick={() => this._spops.updateItem(this.props.context, this.state.project, this._selection.getSelection())
                                    .then((results: ISubmission[]) => {
                                      this.setState({ items: results, showEditform: false });
                                      setTimeout(() => {
                                        //console.log('item group', results);
                                        this._spops.getGrouping(this.state.groupLabels, results)
                                          .then((result: IGroup[]) => {
                                            //console.log('resultttttt group', result);
                                            this.setState({ groups: result });
                                          })
                                      }, 2000);
                                    })}>
                                </PrimaryButton>
                                <DefaultButton text="Cancel" className={'${styles.BtnOutlineTwo} ${styles.mr2}'}
                                  onClick={(e) => { this.setState({ showEditform: false }); }}>
                                </DefaultButton>
                              </div>
                            </div>
                          </div>
                        </div>
                      </Modal>
                      <Modal
                        titleAriaId={titleId}
                        isOpen={this.state.showform}
                        isBlocking={false}>
                        <div className={styles.modal}>
                          <div className="container p-4">
                            <div className={styles.ModalHeader}>
                              <h6>New Submission</h6>
                              <p style={{ color: '#fff' }}>This is the popup about adding Programs,Initiative, Key Achievements in Period,Key Activities for next Period,Support / Attention Needed,Scope , Schedule  etc.</p>
                            </div>
                            <div className={styles.ModalBody}>
                              <div className="dflex mb2">
                                <div className={`{$formgroup} {$mr4}`}>
                                  <Label><strong>Programs :</strong></Label>
                                  <Dropdown
                                    options={this.state.Programs}
                                    placeHolder="Select Program"
                                    defaultSelectedKey="Select Program"
                                    onChange={this.getprogramval}>
                                  </Dropdown>
                                </div>
                                <div className={styles.formgroup}>
                                  <Label><strong>Initiative :</strong></Label>
                                  <Dropdown className={styles.bgone}
                                    defaultValue="Select Initiative"
                                    options={this.state.Initiatives}
                                    onChange={this.getinitiativeval}>
                                  </Dropdown>
                                </div>
                              </div>
                              <div className={styles.formgroup}>
                                <Label><strong>Key Achievements in Period</strong></Label>
                                <TextField className={`{$styles.formcontrol} {$styles.bgone}`}
                                  onChange={this._inputChanged.bind(this)} placeholder="Enter achievements"
                                  name="achievements" required={true} cols={50} rows={5} multiline ></TextField>
                              </div>

                              <div className={styles.formgroup}>
                                <Label><strong>Key Activities for next Period : </strong></Label>
                                <TextField className={`{$styles.formcontrol} {$styles.bgone}`}
                                  onChange={this._inputChanged.bind(this)}
                                  name="activities" required={true} cols={50} rows={5} placeholder="Enter activities" multiline ></TextField>
                              </div>

                              <div className={styles.formgroup}>
                                <Label><strong>Support / Attention Needed : </strong></Label>
                                <TextField className={`{$styles.formcontrol} {$styles.bgone}`}
                                  onChange={this._inputChanged.bind(this)} placeholder="Enter Support needed"
                                  name="supportnd" required={true} cols={50} rows={5} multiline ></TextField>
                              </div>
                              {/* scope container */}
                              <div className="row">
                                <div className="col-6">
                                  <Label><strong>Scope Status :</strong></Label>
                                  <ChoiceGroup defaultSelectedKey={selectedItems.scopsts}
                                    name="scopeStatus"
                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">
                                  <Label><strong>Scope Trend : </strong></Label>
                                  <ChoiceGroup
                                    name="scopeTrend" defaultSelectedKey={selectedItems.scoptrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                              </div>
                              {/* schedule container */}
                              <div className="row">
                                <div className="col-6">
                                  <Label><strong>Schedule Status : </strong></Label>
                                  <ChoiceGroup defaultSelectedKey={selectedItems.schdlsts}
                                    name="scheduleStatus"
                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">
                                  <Label><strong>Schedule Trend : </strong></Label>
                                  <ChoiceGroup
                                    name="scheduleTrend" defaultSelectedKey={selectedItems.schdltrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                              </div>
                              {/* business container */}
                              <div className="row">
                                <div className="col-6">
                                  <Label><strong>Business Case Status :</strong></Label>
                                  <ChoiceGroup
                                    name="businessCaseStatus" defaultSelectedKey={selectedItems.bssts}

                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">
                                  <Label><strong>Business case Trend :</strong></Label>
                                  <ChoiceGroup
                                    name="businessCaseTrend" defaultSelectedKey={selectedItems.bstrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                              </div>
                              {/* overall container */}
                              <div className="row">
                                <div className="col-6">
                                  <Label><strong>Overall Status : </strong></Label>
                                  <ChoiceGroup
                                    name="overallStatus" defaultSelectedKey={selectedItems.ovrlsts}
                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">

                                  <Label><strong>Overall Trend : </strong></Label>
                                  <ChoiceGroup
                                    name="overallTrend" defaultSelectedKey={selectedItems.ovrltrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                              </div>
                              {/* commission container */}
                              <div className="row">
                                <div className="col-6">

                                  <Label><strong>Communication & Comms. Status : </strong></Label>
                                  <ChoiceGroup
                                    name="cncStatus" defaultSelectedKey={selectedItems.cncsts}
                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">
                                  <Label><strong>Communication & Comms. Trend : </strong></Label>
                                  <ChoiceGroup
                                    name="cncTrend" defaultSelectedKey={selectedItems.cnctrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />

                                </div>
                              </div>
                              {/* impact container */}
                              <div className="row">
                                <div className="col-6">
                                  <Label><strong>Impact on opt Status :</strong></Label>
                                  <ChoiceGroup
                                    name="imOpStatus" defaultSelectedKey={selectedItems.imopsts}
                                    options={[
                                      {
                                        key: 'On schedule;goals within reach',
                                        text: 'On schedule;goals within reach'
                                      },
                                      {
                                        key: 'minor issues threatening scheduleand/or goals',
                                        text: 'minor issues threatening scheduleand/or goals'
                                      },
                                      {
                                        key: 'Behind schedule and/or goals are risk',
                                        text: 'Behind schedule and/or goals are risk'
                                      }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                                <div className="col-6">
                                  <Label><strong>Impact on opt Trend : </strong></Label>
                                  <ChoiceGroup
                                    name="imOpTrend" defaultSelectedKey={selectedItems.imoptrnd}
                                    options={[
                                      { key: 'Trending up', text: 'Trending up' },
                                      { key: 'Trending down', text: 'Trending down' },
                                      { key: 'Stable', text: 'Stable' }
                                    ]}
                                    onChange={this._handlebtnchange}
                                  />
                                </div>
                              </div>
                              <div className={styles.ModalFooter}>
                                <PrimaryButton text="Save" type="submit" className={styles.BtneOne}
                                  onClick={() => this._spops.createItem(this.props.context, this.state.project)
                                    .then((results: ISubmission[]) => {
                                      this.setState({ items: results, showform: false });
                                      setTimeout(() => {
                                        //console.log('item group', results);
                                        this._spops.getGrouping(this.state.groupLabels, results)
                                          .then((result: IGroup[]) => {
                                            //console.log('resultttttt group', result);
                                            this.setState({ groups: result });
                                          })
                                      }, 2000);
                                    })}>
                                </PrimaryButton>
                                <DefaultButton text="Cancel" className={'${styles.BtnOutlineTwo} ${styles.mr2}'}
                                  onClick={(e) => { this.setState({ showform: false }); }}>
                                </DefaultButton>
                              </div>
                            </div>
                          </div>
                        </div>
                      </Modal>
                    </div>
                    <div className={styles.ReportsSection}>
                      <div className={styles.StatsContainer}>
                        <div className={this.state.elementId == '1'?styles.statsCard+' '+styles.active:styles.statsCard} onClick={() => this._initiativeTopFilter('All','1')}>
                          <h6>Total Submissions</h6>
                          <h1 className={styles.countone}>{this.state.allCount}</h1>
                        </div>
                        <div className={this.state.elementId == '2'?styles.statsCard+' '+styles.active:styles.statsCard} onClick={() => this._initiativeTopFilter('','2')}>
                          <h6>My Submissions</h6>
                          <h1 className={styles.counttwo}>{this.state.userCount}</h1>
                        </div>
                      </div>
                      <div className="d-none">
                        <label className={styles.container1}>On Schedule
                         <input type="radio" name="radio" />
                          <span className={styles.checkmark}></span>
                        </label>
                        <label className={styles.container2}>Minor Issues
                         <input type="radio" name="radio" />
                          <span className={styles.checkmark}></span>
                        </label>
                        <label className={styles.container3}>Need Help
                        <input type="radio" name="radio" />
                          <span className={styles.checkmark}></span>
                        </label>
                      </div>
                    </div>
                    <div className="row m-2">
                      <div className="col-6">
                        <div className="row">
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '3'?'card card-body '+styles.active:'card card-body'} style={{margin:'3px'}} id="ini_3" onClick={() => { this._initiativeScopeFilter('All', 'On schedule;goals within reach','3') }}>
                              <p>On Schedule</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.All.OnSchedule}</span>
                                <span><i style={{ fontSize: '30px' }} className="fa fa-calendar " aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '4'?'card card-body '+styles.active:'card card-body'} style={{margin:'3px'}} id="ini_4" onClick={() => { this._initiativeScopeFilter('All', 'minor issues threatening scheduleand/or goals','4') }}>
                              <p>Minor Issues</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.All.MinurIssues}</span>
                                <span><i style={{ fontSize: '30px' }} className="fas fa-exclamation-triangle" aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '5'?'card card-body '+styles.active:'card card-body'} style={{margin:'3px'}} id="ini_5" onClick={() => { this._initiativeScopeFilter('All', 'Behind schedule and/or goals are risk','5') }}>
                              <p>Need Help</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.All.NeedHelp}</span>
                                <span><i style={{ fontSize: '30px' }} className="fas fa-hands-helping" aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                        </div>

                        <div className="row">
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '6'?'card card-body '+styles.active:'card card-body'} style={{margin:'3px',color: '#bceb3c'}} id="ini_6" onClick={() => { this._initiativeTrendFilter('All', 'Trending up','6') }}>
                              <p>Trending Up</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.All.TrendingUp}</span>
                                <span><i style={{ fontSize: '30px' }} className="fas fa-arrow-circle-up" aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '7'?'card card-body '+styles.active:'card card-body'} id="ini_7" style={{ color: '#f5a31a' ,margin:'3px'}} onClick={() => { this._initiativeTrendFilter('All', 'Stable','7') }}>
                              <p>Stable</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.All.Stable}</span>
                                <span><i style={{ fontSize: '30px' }} className="fas fa-arrow-circle-right" aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '8'?'card card-body '+styles.active:'card card-body'} id="ini_8" style={{ color: '#f05d23',margin:'3px' }} onClick={() => { this._initiativeTrendFilter('All', 'Trending down','8') }}>
                              <p>Trening Down</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.All.TreningDown}</span>
                                <span><i style={{ fontSize: '30px' }} className="fas fa-arrow-circle-down" aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                        </div>

                      </div>
                      <div className="col-6">
                        <div className="row ml-0">
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '9'?'card card-body '+styles.active:'card card-body'} style={{margin:'3px'}} id="ini_9" onClick={() => { this._initiativeScopeFilter('', 'On schedule;goals within reach','9') }}>
                              <p>On Schedule</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.User.OnSchedule}</span>
                                <span><i style={{ fontSize: '30px' }} className="fa fa-calendar " aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '10'?'card card-body '+styles.active:'card card-body'} style={{margin:'3px'}} id="ini_10" onClick={() => { this._initiativeScopeFilter('', 'minor issues threatening scheduleand/or goals','10') }}>
                              <p>Minor Issues</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.User.MinurIssues}</span>
                                <span><i style={{ fontSize: '30px' }} className="fas fa-exclamation-triangle" aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '11'?'card card-body '+styles.active:'card card-body'} style={{margin:'3px'}} id="ini_11" onClick={() => { this._initiativeScopeFilter('', 'Behind schedule and/or goals are risk','11') }}>
                              <p>Need Help</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.User.NeedHelp}</span>
                                <span><i style={{ fontSize: '30px' }} className="fas fa-hands-helping" aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                        </div>

                        <div className="row ml-0">
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '12'?'card card-body '+styles.active:'card card-body'} id="ini_12" style={{ color: '#bceb3c' ,margin:'3px'}} onClick={() => { this._initiativeTrendFilter('', 'Trending up','12') }}>
                              <p>Trending Up</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.User.TrendingUp}</span>
                                <span><i style={{ fontSize: '30px' }} className="fas fa-arrow-circle-up" aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '13'?'card card-body '+styles.active:'card card-body'}  id="ini_13" style={{ color: '#f5a31a',margin:'3px' }} onClick={() => { this._initiativeTrendFilter('', 'Stable','13') }}>
                              <p>Stable</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.User.Stable}</span>
                                <span><i style={{ fontSize: '30px' }} className="fas fa-arrow-circle-right" aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                          <div className="col-4 p-0">
                            <div className={this.state.elementId == '14'?'card card-body '+styles.active:'card card-body'} id="ini_14" style={{ color: '#f05d23',margin:'3px' }} onClick={() => { this._initiativeTrendFilter('', 'Trending down','14') }}>
                              <p>Trening Down</p>
                              <h3 className="d-flex justify-content-between">
                                <span>{this.state.dashboardCounts.User.TreningDown}</span>
                                <span><i style={{ fontSize: '30px' }} className="fas fa-arrow-circle-down" aria-hidden="true"></i></span>
                              </h3>
                            </div>
                          </div>
                        </div>

                      </div>

                    </div>

                    {/* <Label>Progress Report</Label> */}
                    {/* <div >{selectionDetails}</div> */}
                    <br />
                    <Announced message={selectionDetails} />
                    {announcedMessage ? <Announced message={announcedMessage} /> : undefined}
                    {isModalSelection ? (
                      <MarqueeSelection selection={this._selection}>
                        <DetailsList
                          items={items}
                          groups={groups}
                          columns={columns}
                          selectionMode={SelectionMode.multiple}
                          //getKey={this._getKey}
                          setKey="multiple"
                          layoutMode={DetailsListLayoutMode.justified}
                          isHeaderVisible={true}
                          selection={this._selection}
                          selectionPreservedOnEmptyClick={true}
                          //onItemInvoked={this._onItemInvoked}
                          enterModalSelectionOnTouch={true}
                          onRenderDetailsHeader={this._onRenderDetailsHeader}
                          ariaLabelForSelectionColumn="Toggle selection"
                          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                          checkButtonAriaLabel="Row checkbox"
                          onRenderItemColumn={this._onRenderColumn}
                          groupProps={{
                            showEmptyGroups: true,
                          }}
                        />
                      </MarqueeSelection>
                    ) : (
                        <DetailsList
                          items={items}
                          columns={columns}
                          groups={groups}
                          selectionMode={SelectionMode.none}
                          //getKey={this._getKey}
                          setKey="none"
                          layoutMode={DetailsListLayoutMode.justified}
                          isHeaderVisible={true}
                          onRenderDetailsHeader={this._onRenderDetailsHeader}
                          onRenderItemColumn={this._onRenderColumn}
                          groupProps={{
                            showEmptyGroups: true,
                          }}
                        />
                      )}
                  </Fabric>
                </div>
              </Label>
            </PivotItem>
            <PivotItem headerText="Submission & Analysis" itemIcon="AnalyticsView">
              <Label styles={labelStyles} className="my-2">
                <Dashboard itemcount={this.state.allCount} userCount={this.state.userCount} Items={this.state.items} Programs={this.state.Programs} {...this.props} />
              </Label>
            </PivotItem>
            <PivotItem headerText="Need Attention" itemIcon="FunctionalManagerDashboard" >
              <Label styles={labelStyles} className="my-2">
                <NeedAttention Programs={this.state.Programs} {...this.props} />
              </Label>
            </PivotItem>
            <PivotItem headerText="Dashboard" itemIcon="FunctionalManagerDashboard" >
              <Label styles={labelStyles} className="my-2">
                <Reports Programs={this.state.Programs} {...this.props} />
              </Label>
            </PivotItem>
          </Pivot>
        </div>
      </div>
    );
  }
}
