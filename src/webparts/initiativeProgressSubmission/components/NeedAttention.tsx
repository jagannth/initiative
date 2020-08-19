import * as React from 'react';
import { IInitiativeProgressSubmissionProps } from '../models/IInitiativeProgressSubmissionProps';
import { Crudoperations } from '../services/SPServices';
import { IReportState } from '../models/IReportState';

import styles from './styles.module.scss';
import { Label, Dropdown, Icon, PrimaryButton } from 'office-ui-fabric-react';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import pptxgen from "pptxgenjs";
import { INeedAttentionState } from '../models/IDashboardState';

// Used to add spacing between example checkboxes
const stackTokens = { childrenGap: 10 };

export class NeedAttention extends React.Component<IInitiativeProgressSubmissionProps, INeedAttentionState, {}>{
    public _spops: Crudoperations;

    constructor(props: IInitiativeProgressSubmissionProps) {
        super(props);
        this._spops = new Crudoperations();
        this.state = {
            needAttention: []
        }
    }
    public componentDidMount() {
        this._spops.needAttentionList(this.props.context)
            .then((result) => {
                setTimeout(() => {
                    console.log('need attention', result);
                    this.setState({
                        needAttention: result
                    })
                }, 1000);
            });
    }
    public render(): React.ReactElement {

        return (<div>
            <div className="row">
                <div className="card card-body" style={{ border: 'none', height: '100%' }}>

                    <div>
                        <strong className={styles.counttwo}>Need Attention</strong>
                        <Dropdown
                            options={this.props.Programs}
                            placeHolder="Select Program">
                        </Dropdown>
                        <hr />
                        {this.state.needAttention.map((x: any) => {
                            console.log('working');
                            return (<div className="row m-2">
                                {/* <div className="col-1">
                                    <img src={require("../images/alert.png")} style={{ width: '20px' }} />
                                </div> */}
                                <div className="col-12">
                                    <h5>Program : {x.Programs}</h5>
                                    <h6>Initiative : {x.Initiative} {x.OveralTrend}</h6>
                                    <p><strong><img src={require("../images/alert.png")} style={{ width: '17px', marginRight: '7px' }} />Help Text :</strong> {x.supportattentionneeded}</p>
                                </div>
                            </div>)
                        })}
                    </div>
                </div>
            </div>
        </div>)
    }
}