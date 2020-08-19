import * as React from 'react';
import { IInitiativeProgressSubmissionProps } from '../models/IInitiativeProgressSubmissionProps';
import { Crudoperations } from '../services/SPServices';
import { IReportState } from '../models/IReportState';

import styles from './styles.module.scss';
import { Label, Dropdown, Icon, PrimaryButton } from 'office-ui-fabric-react';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import pptxgen from "pptxgenjs";

// Used to add spacing between example checkboxes
const stackTokens = { childrenGap: 10 };

export class Reports extends React.Component<IInitiativeProgressSubmissionProps, IReportState, {}>{
    public _spops: Crudoperations;

    constructor(props: IInitiativeProgressSubmissionProps) {
        super(props);
        this._spops = new Crudoperations();
        this.state = {
            initiativeList: [],
            allInitiativeList: [],
            isAll: true,
            selectedItem: this._getselectedItem(),
            isInitiative: false,
            selectedInitiative: {}
        }
        this.generateReport = this.generateReport.bind(this);
    }


    public render(): React.ReactElement {

        return (
            <div className={styles.initiativeProgressSubmission}>
                <div className={styles.padd}>
                    <div className={styles.contianer}>
                        <div className="row">
                            <div className="col-12 pr-0">
                                <div className={styles.ReportsSection}>
                                    <div className={styles.ProgramStatsCard}>
                                        <div className={styles.CardHeader}>
                                            <h6>Program Dashboard</h6>
                                        </div>
                                        <div className="card card-body">
                                            {this.state.isInitiative && <a className="text-primary" onClick={() => { this.setState({ isInitiative: false }) }}>	&lt; Back to Program Dashboard Summery</a>}
                                            <div className="row">
                                                <div className="col-6">
                                                    <Label>Programs:</Label>
                                                    <Dropdown
                                                        options={this.props.Programs}
                                                        placeHolder="Select Program"
                                                        defaultSelectedKey={this.state.selectedItem}
                                                        onChange={this.generateReport}>
                                                    </Dropdown>
                                                    <div className="mt-2">
                                                        {/* <PrimaryButton className={styles.BtneOne} text="Generate Report"></PrimaryButton> */}
                                                        {/* <PrimaryButton className={styles.BtneOne} style={{ marginRight: '5px' }} text="Generate PDF"></PrimaryButton>
                                                        <PrimaryButton className={styles.BtneOne} text="Generate PPT" onClick={this.generatePPT}></PrimaryButton> */}
                                                    </div>
                                                </div>
                                                {this.state.isInitiative &&
                                                    <div className="col-6">
                                                        <Label>Initiative:</Label>
                                                        <select className="form-control" onChange={this.getInitiative1}>
                                                            <option disabled>--select --</option>
                                                            {this.state.initiativeList.map((x: any) => {
                                                                return <option value={x.Initiative}>{x.Initiative}</option>
                                                            })}
                                                        </select>
                                                    </div>}
                                                <div className="col-9 mt-4">
                                                    <PrimaryButton className={styles.BtneOne} style={{ marginRight: '5px' }} text="Generate PDF"></PrimaryButton>
                                                    <PrimaryButton className={styles.BtneOne} style={{ marginRight: '5px' }} text="Generate PPT" onClick={this.generatePPT}></PrimaryButton>
                                                    {this.state.isInitiative && <PrimaryButton className={styles.BtneOne} text="Reset"></PrimaryButton>}
                                                    {/* <Stack tokens={stackTokens}>
                                                        <Checkbox label="Print Program Dashboard Summery" defaultChecked={true} />
                                                        <Checkbox label="Print Program Dashboard" defaultChecked={true} />
                                                        <Checkbox label="Print Program Dashboard Details" defaultChecked={true} />
                                                    </Stack> */}
                                                </div>
                                            </div><hr />
                                            <div>
                                                <div className="mb-2">
                                                    <span>
                                                        <span className="px-2">
                                                            <i className="fas fa-stop mr-1" style={{ color: '#bceb3c' }}></i>On Schedule
                                                    </span>
                                                        <span className="px-2">
                                                            <i className="fas fa-stop mr-1" style={{ color: '#f5a31a' }}></i>Minor Issues
                                                    </span>
                                                        <span className="px-2">
                                                            <i className="fas fa-stop mr-1" style={{ color: '#f05d23' }}></i>Need Help
                                                    </span>
                                                    </span>
                                                    <span className="float-right">
                                                        <span className="px-2">
                                                            <i className="fas fa-play mr-1" style={{ color: 'gray', transform: 'rotate(-90deg)' }}></i>Trening Up
                                                    </span>
                                                        <span className="px-2">
                                                            <i className="fas fa-play mr-1" style={{ color: 'gray' }}></i>Stable
                                                    </span>
                                                        <span className="px-2">
                                                            <i className="fas fa-play mr-1" style={{ color: 'gray', transform: 'rotate(90deg)' }}></i>Trending Down
                                                    </span>
                                                    </span>
                                                </div>
                                                {this.state.isInitiative && <div>
                                                    <div className="row">
                                                        <div className="col-6">
                                                            <h5>Key achievements in Period</h5>
                                                            <div style={{ border: '1px solid black', height:'140px' }}>
                                                                {this.state.selectedInitiative.Initiative[0].keyachievementsinperiod}
                                                            </div>
                                                        </div>
                                                        <div className="col-6">
                                                            <h5>Key activities in Next Period</h5>
                                                            <div style={{ border: '1px solid black', height:'140px' }}>
                                                                {this.state.selectedInitiative.Initiative[0].keyactivitiesfornextperiod}
                                                            </div>
                                                        </div>
                                                    </div>
                                                    <div className="row">
                                                        <div className="col-12">
                                                            <h5>Support / Attention Needed</h5>
                                                            <div style={{ border: '1px solid black', height:'60px' }}>
                                                                {this.state.selectedInitiative.Initiative[0].supportattentionneeded}
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>}
                                                {this.state.isInitiative == false && <table className="table table-bordered" id="tabAutoPaging">
                                                    <thead className="text-center">
                                                        <tr className="bg-light">
                                                            <th scope="col" style={{ width: '25%' }}>Programs</th>
                                                            <th scope="col" style={{ width: '25%' }}>Initiative</th>
                                                            <th scope="col" style={{ width: '9%' }}>Scope</th>
                                                            <th scope="col" style={{ width: '9%' }}>Schedule</th>
                                                            <th scope="col" style={{ width: '9%' }}>Business</th>
                                                            <th scope="col" style={{ width: '9%' }}>Overall</th>
                                                            <th scope="col" style={{ width: '14%' }}>Report Date</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody className="text-center">
                                                        {this.state.initiativeList.map((x: any) => {
                                                            return (
                                                                <tr>
                                                                    {x.Status ? <td className="align-middle" rowSpan={x.Count}>{x.Programs}</td> : ''}
                                                                    <td className="align-middle" onClick={() => { this.getInitiative(x.Initiative) }}>{x.Initiative} </td>
                                                                    {(() => {
                                                                        let colorr = '#a6cb12';
                                                                        if (x.ScopeStatus == 'minor issues threatening scheduleand/or goals') colorr = '#f5a31a';
                                                                        if (x.ScopeStatus == 'Behind schedule and/or goals are risk') colorr = '#f05d23';
                                                                        switch (x.ScopeTrend) {
                                                                            case 'Trending up':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play" style={{ transform: 'rotate(-90deg)' }}></i></td>
                                                                            case 'Trending down':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play" style={{ transform: 'rotate(90deg)' }}></i></td>
                                                                            case 'Stable':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play"></i></td>
                                                                            default:
                                                                                return <td className="align-middle"></td>
                                                                        }
                                                                    })()}
                                                                    {(() => {
                                                                        let colorr = '#a6cb12';
                                                                        if (x.ScheduleStatus == 'minor issues threatening scheduleand/or goals') colorr = '#f5a31a';
                                                                        if (x.ScheduleStatus == 'Behind schedule and/or goals are risk') colorr = '#f05d23';
                                                                        switch (x.ScheduleTrend) {
                                                                            case 'Trending up':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play" style={{ transform: 'rotate(-90deg)' }}></i></td>
                                                                            case 'Trending down':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play" style={{ transform: 'rotate(90deg)' }}></i></td>
                                                                            case 'Stable':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play"></i></td>
                                                                            default:
                                                                                return <td className="align-middle"></td>
                                                                        }
                                                                    })()}
                                                                    {(() => {
                                                                        let colorr = '#a6cb12';
                                                                        if (x.BusinessCaseStatus == 'minor issues threatening scheduleand/or goals') colorr = '#f5a31a';
                                                                        if (x.BusinessCaseStatus == 'Behind schedule and/or goals are risk') colorr = '#f05d23';
                                                                        switch (x.BusinessCaseTrend) {
                                                                            case 'Trending up':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play" style={{ transform: 'rotate(-90deg)' }}></i></td>
                                                                            case 'Trending down':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play" style={{ transform: 'rotate(90deg)' }}></i></td>
                                                                            case 'Stable':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play"></i></td>
                                                                            default:
                                                                                return <td className="align-middle"></td>
                                                                        }
                                                                    })()}
                                                                    {(() => {
                                                                        let colorr = '#a6cb12';
                                                                        if (x.OverallStatus == 'minor issues threatening scheduleand/or goals') colorr = '#f5a31a';
                                                                        if (x.OverallStatus == 'Behind schedule and/or goals are risk') colorr = '#f05d23';
                                                                        switch (x.OveralTrend) {
                                                                            case 'Trending up':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play" style={{ transform: 'rotate(-90deg)' }}></i></td>
                                                                            case 'Trending down':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play" style={{ transform: 'rotate(90deg)' }}></i></td>
                                                                            case 'Stable':
                                                                                return <td className="align-middle" style={{ backgroundColor: colorr, color: 'White' }}><i className="fas fa-play"></i></td>
                                                                            default:
                                                                                return <td className="align-middle"></td>
                                                                        }
                                                                    })()}
                                                                    <td className="align-middle">{x.Created}</td>
                                                                </tr>
                                                            )
                                                        })}

                                                    </tbody>
                                                </table>
                                                } </div>
                                            <div>
                                                {/* <img src={require('../images/novartis-logo-preview-image.png')} alt="test" /> */}
                                            </div>
                                        </div>
                                    </div>
                                </div>

                            </div>
                        </div>
                    </div>

                </div>
            </div>

        )
    }
    private _getselectedItem(): string {
        // const selectionText = this.props.Programs[0].text;
        setTimeout(() => {
            this._spops.getAllReports(this.props.context, this.props.Programs)
                .then((result: any) => {
                    setTimeout(() => {
                        console.log('allInitiativeList list', result);
                        this.setState({
                            initiativeList: result
                        });
                    }, 2000);
                });
        }, 1000);
        return 'Select Program';
    }
    getInitiative(value) {
        console.log('value',value);
        this._spops.getTrends(this.props.context, value, 1).then((result: any) => {
            console.log('get initiative', result);
            setTimeout(() => {
                this.setState({
                    selectedInitiative: result,
                    isInitiative: true
                })
            }, 600);
        })
    }
    getInitiative1 = (event) => {
        this._spops.getTrends(this.props.context, event.target.value, 1).then((result: any) => {
            console.log('get initiative1', result);
            setTimeout(() => {
                this.setState({
                    selectedInitiative: result,
                    isInitiative: true
                })
            }, 600);
        })
    }
    public generateReport = (event: any, result: any) => {
        let prgm = result.text;
        console.log('result.text test', result.text);
        this._spops.getReports(this.props.context, prgm)
            .then((result: any) => {
                console.log('initiative list', result);
                this.setState({
                    initiativeList: result
                });
            });
    }
    public generatePPT = () => {
        // // 1. Create a new Presentation
        let pptx = new pptxgen();
        pptx.defineLayout({ name: 'A3', width: 22, height: 15 });
        // // 2. Add a Slide
        let slide = pptx.addSlide('slideone');
        let textboxOpts = { x: 0.5, y: 0, color: "363636", fontSize: 20, fontFace: 'Arial', bold: true, align: pptx.AlignH.left };
        let textboxText = "";
        var rows = [];
        rows.push([{ text: 'Initiative', options: { align: "center", bold: true } }, { text: 'Scope', options: { align: "center", bold: true } }, { text: 'Schedule', options: { align: "center", bold: true } }, { text: 'Business', options: { align: "center", bold: true } }, { text: 'Overall Status', options: { align: "center", bold: true } }, { text: 'Last Reported Date', options: { align: "center", bold: true } }]);
        this.state.initiativeList.map((x: any) => {
            textboxText = x.Programs;
            let clr1 = 'a6cb12'; let clr2 = 'a6cb12'; let clr3 = 'a6cb12'; let clr4 = 'a6cb12';
            switch (x.ScopeStatus) {
                case 'Behind schedule and/or goals are risk': clr1 = 'f05d23';
                case 'minor issues threatening scheduleand/or goals': clr1 = 'f5a31a';
            }
            switch (x.ScheduleStatus) {
                case 'Behind schedule and/or goals are risk': clr2 = 'f05d23';
                case 'minor issues threatening scheduleand/or goals': clr2 = 'f5a31a';
            }
            switch (x.BusinessCaseStatus) {
                case 'Behind schedule and/or goals are risk': clr3 = 'f05d23';
                case 'minor issues threatening scheduleand/or goals': clr3 = 'f5a31a';
            }
            switch (x.OveralStatus) {
                case 'Behind schedule and/or goals are risk': clr4 = 'f05d23';
                case 'minor issues threatening scheduleand/or goals': clr4 = 'f5a31a';
            }

            rows.push([
                { text: x.Initiative, options: { w: 4 } },
                { text: '▶', options: { Color: "f5a31a", fill: clr1, valign: "center", align: "center" } },
                { text: '▶', options: { Color: "f5a31a", fill: clr2, valign: "center", align: "center" } },
                { text: '▶', options: { Color: "f5a31a", fill: clr3, valign: "center", align: "center" } },
                { text: '▶', options: { Color: "f5a31a", fill: clr4, valign: "center", align: "center" } },
                { text: x.Created, options: { Color: "f5a31a", valign: "center", align: "center" } },
            ]);
        })
        slide.addText(textboxText, textboxOpts);
        slide.addTable(rows, { x: 0.5, y: 0.5, w: 10.0, colW: [3.0, 1.0, 1.0, 1.0, 1.0, 2.0] });
        slide.addImage({ path: require('../images/novartis-logo-preview-image.png'), x: 6.5, y: 5, w: 3.0, h: 0.5 });
        this.state.initiativeList.map((x: any) => {
            this.testMethod_Table(pptx, x, textboxOpts);
        })

        // this.testMethod_Chart(pptx);
        pptx.writeFile("progres-report.pptx");
    }
    testMethod_Chart(pptx: pptxgen) {
        let slide = pptx.addSlide("MASTER_SLIDE chart");

        let dataChart = [
            {
                name: "Region 1",
                labels: ["May", "June", "July", "August", "September"],
                values: [26, 53, 100, 75, 41],
            },
        ];
        slide.addChart(pptx.ChartType.bar, dataChart, { x: 0.5, y: 2.5, w: 5.25, h: 4 }); // TEST: charts
    }
    testMethod_Table(pptx: pptxgen, item: any, options: any) {
        const slidee = pptx.addSlide('slideone');
        slidee.addText(item.Initiative, options);
        const rows = [];
        const rows2 = [];
        const rows3 = [];
        let clr1 = 'a6cb12'; let clr2 = 'a6cb12'; let clr3 = 'a6cb12'; let clr4 = 'a6cb12';
        switch (item.ScopeStatus) {
            case 'Behind schedule and/or goals are risk': clr1 = 'f05d23';
            case 'minor issues threatening scheduleand/or goals': clr1 = 'f5a31a';
        }
        switch (item.ScheduleStatus) {
            case 'Behind schedule and/or goals are risk': clr2 = 'f05d23';
            case 'minor issues threatening scheduleand/or goals': clr2 = 'f5a31a';
        }
        switch (item.BusinessCaseStatus) {
            case 'Behind schedule and/or goals are risk': clr3 = 'f05d23';
            case 'minor issues threatening scheduleand/or goals': clr3 = 'f5a31a';
        }
        switch (item.OveralStatus) {
            case 'Behind schedule and/or goals are risk': clr4 = 'f05d23';
            case 'minor issues threatening scheduleand/or goals': clr4 = 'f5a31a';
        }
        rows.push([{ text: "Scope", options: { align: "center", bold: true } }, { text: "Schedule", options: { align: "center", bold: true } }, { text: "Chang & Comms", options: { align: "center", bold: true } }, { text: "Overall Status", options: { align: "center", bold: true } }]);
        rows2.push([{ text: "Key Achievements in Period", options: { fontSize: 15, fontFace: 'Arial', bold: true } }, { text: "Key Activities in Next Period", options: { fontSize: 15, fontFace: 'Arial', bold: true } }]);
        rows3.push([{ text: "Support / Attention Needed", options: { fontSize: 15, fontFace: 'Arial', bold: true } }]);
        rows.push([
            { text: '▶', options: { Color: "f5a31a", fill: clr1, align: "center" } },
            { text: '▶', options: { Color: "a6cb12", fill: clr2, align: "center" } },
            { text: '▶', options: { Color: "d32626", fill: clr3, align: "center" } },
            { text: '▶', options: { Color: "a6cb12", fill: clr4, align: "center" } },
        ]);
        rows2.push([
            { text: item.keyachievementsinperiod, options: { Color: "f5a31a", h: 2.0 } },
            { text: item.keyactivitiesfornextperiod, options: { Color: "a6cb12", h: 2.0 } }
        ]);
        rows3.push([
            { text: item.supportattentionneeded, options: { Color: "f5a31a" } }
        ]);
        slidee.addTable(rows, { x: 5.5, y: 0.0, w: 4, h: 1 });
        slidee.addTable(rows2, { x: 0.5, y: 1.5, h: 2, color: "363636" });
        slidee.addTable(rows3, { x: 0.5, y: 4.0, color: "363636" });
        slidee.addImage({ path: require('../images/novartis-logo-preview-image.png'), x: 6.5, y: 5, w: 3.0, h: 0.5 });
    }
    public componentDidMount() {
        console.log('programs inititatives', this.props.Programs);
    }
}