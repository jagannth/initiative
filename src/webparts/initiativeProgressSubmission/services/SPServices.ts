import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IDropdownOption, IObjectWithKey, IGroup } from "office-ui-fabric-react";
import { ISubmission } from "../models/ISubmission";
import { Analysis } from '../models/Analysis';
import { Provider } from "react";
import { Category, TrendAnalysis } from "../models/Category";
import * as moment from 'moment';



export class Crudoperations {
    public analysis: Analysis;
    constructor() {
        this.analysis = new Analysis();
    }


    public getprograms(context: WebPartContext, _listname: string): Promise<IDropdownOption[]> {
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('" + _listname + "')/items?$select=Title";
        // let geturl:string="https://sticsoftio.sharepoint.com/sites/POC/_api/web/lists/GetByTitle('Datacollect')/fields?$filter=EntityPropertyName eq '" +_listname+"'";
        var listTitles: IDropdownOption[] = [];
        let choicecount: number;
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            context.spHttpClient
                .get(geturl, SPHttpClient.configurations.v1).then(
                    (Response: SPHttpClientResponse) => {
                        Response.json().then((results: any) => {
                            results.value.map((result: any) => {
                                listTitles.push({
                                    key: result.Title, text: result.Title,
                                });
                            });
                        });
                        resolve(listTitles);
                    }, (error: any): void => {
                        reject("error ocuured" + error);
                    }
                );
        });
    }

    public getinitiative(context: WebPartContext, _program: string): Promise<IDropdownOption[]> {
        //let geturl:string=context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('Initiatives')/items?$select=Title";
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Initiatives')/items?$select=Title,Bundle/Title&$expand=Bundle&$filter=Bundle/Title eq '" + _program + "'";
        // let geturl:string="https://sticsoftio.sharepoint.com/sites/POC/_api/web/lists/GetByTitle('Datacollect')/fields?$filter=EntityPropertyName eq '" +_listname+"'";
        var listTitles: IDropdownOption[] = [];
        let choicecount: number;
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            context.spHttpClient
                .get(geturl, SPHttpClient.configurations.v1).then(
                    (Response: SPHttpClientResponse) => {
                        Response.json().then((results: any) => {
                            results.value.map((result: any) => {
                                //result.Choices.map((choice:any)=>{
                                listTitles.push({
                                    //key:choice,text:choice, 
                                    key: result.Title, text: result.Title,
                                });
                                //});                         

                            });
                        });
                        resolve(listTitles);
                    }, (error: any): void => {
                        reject("error ocuured" + error);
                    }
                );
        });
    }

    public getChoicesone(context: WebPartContext): Promise<IDropdownOption[]> {
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/POC/_api/web/lists/GetByTitle('ChoiceMaster')/items?$select=Title";
        var firstchoices: IDropdownOption[] = [];
        let choicecount: number;
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            context.spHttpClient
                .get(geturl, SPHttpClient.configurations.v1).then(
                    (Response: SPHttpClientResponse) => {
                        Response.json().then((results: any) => {
                            results.value.map((result: any) => {
                                firstchoices.push({
                                    key: result.Title, text: result.Title,
                                });

                            });
                        });
                        resolve(firstchoices);
                    }, (error: any): void => {
                        reject("error ocuured" + error);
                    }
                );
        });
    }

    public getGrouping1(context: WebPartContext): Promise<IGroup[]> {
        //let geturl:string=context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('Datacollect')/Items?$orderby=Initiative asc";
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/Items?";
        geturl = geturl + "&select=*";
        geturl = geturl + "&$orderby=ID desc";
        //let geturl:string=context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('Datacollect')/Items?$select=Initiative&$filter=Initiative eq'"+Group+"'";
        var Groups: IGroup[] = [];
        const _groupfield = [];
        let itemcount: number;
        let prevcount = 0;
        const val = new Promise<any>(async (resolve, reject) => {
            context.spHttpClient
                .get(geturl, SPHttpClient.configurations.v1).then(
                    (Response: SPHttpClientResponse) => {
                        Response.json().then((results: any) => {
                            results.value.map((result: any) => {
                                itemcount = results.value.length;
                                if (_groupfield.indexOf(result.Initiative) === -1) {
                                    console.log('inside main result', result);
                                    _groupfield.push(result.Initiative);
                                    itemcount = results.value.length;
                                }
                            })
                        });
                        resolve(Groups);
                    }, (error: any): void => {
                        reject("error ocuured" + error);
                    }
                );
        });
        console.log('real val', val);
        return val;
    }
    public async getGrouping(groupList: string[], items: ISubmission[]): Promise<IGroup[]> {
        var Groups: IGroup[] = [];
        let itemcount: number;
        let prevcount = 0;
        groupList.map((x, index) => {
            itemcount = items.filter(z => z.intiate == x).length;
            Groups.push({
                key: "group" + index,
                name: x,
                startIndex: prevcount,
                count: itemcount
            })
            prevcount = prevcount + itemcount
        });
        return Groups;
    }

    public async getGroupingLable(context: WebPartContext): Promise<string[]> {
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/Items?";
        geturl = geturl + "&select=Initiative";
        geturl = geturl + "&$orderby=Initiative";
        const _groupfield = [];
        await context.spHttpClient
            .get(geturl, SPHttpClient.configurations.v1).then(
                (Response: SPHttpClientResponse) => {
                    Response.json().then((results: any) => {
                        results.value.map((result: any) => {
                            //itemcount=results.value.length;
                            if (_groupfield.indexOf(result.Initiative) === -1) {
                                _groupfield.push(result.Initiative);
                            }
                        });
                    });
                }).catch(error => {
                    console.log('error', error);
                });
        return _groupfield;
    }


    public createItem(context: WebPartContext, _listinitiate: ISubmission): Promise<ISubmission[]> {
        let posturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/items";

        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/Items?";
        geturl = geturl + "&select=*";
        geturl = geturl + "&$orderby=Initiative";

        var items: ISubmission[] = [];
        var close: boolean = false;
        const body: string = JSON.stringify({
            Programs: _listinitiate.program,
            Initiative: _listinitiate.intiate,
            keyachievementsinperiod: _listinitiate.achievements,
            keyactivitiesfornextperiod: _listinitiate.activities,
            supportattentionneeded: _listinitiate.supportnd,
            ScopeStatus: _listinitiate.scopsts,
            ScheduleStatus: _listinitiate.schdlsts,
            BusinessCaseStatus: _listinitiate.bssts,
            OverallStatus: _listinitiate.ovrlsts,
            ScopeTrend: _listinitiate.scoptrnd,
            ScheduleTrend: _listinitiate.schdltrnd,
            BusinessCaseTrend: _listinitiate.bstrnd,
            OveralTrend: _listinitiate.ovrltrnd,
            ChangecommasStatus: _listinitiate.cncsts,
            ChangecommasTrend: _listinitiate.cnctrnd,
            ImpactonOperationsStatus: _listinitiate.imopsts,
            ImpactonOperationsTrend: _listinitiate.imoptrnd,

        });
        const options: ISPHttpClientOptions = {
            headers: {
                Accept: "application/json;odata=nometadata",
                "content-type": "application/json;odata=nometadat",
                "odataverion": ""
            },
            body: body,
        };
        return new Promise<ISubmission[]>(async (resolve, reject) => {
            context.spHttpClient.post(posturl, SPHttpClient.configurations.v1, options).then(
                () => {
                    context.spHttpClient.get(geturl, SPHttpClient.configurations.v1).then(
                        (Response: SPHttpClientResponse) => {
                            Response.json().then((results: any) => {
                                results.value.map((result: any) => {
                                    items.push({
                                        program: result.Programs,
                                        intiate: result.Initiative,
                                        scopsts: result.ScopeStatus,
                                        schdlsts: result.ScheduleStatus,
                                        bssts: result.BusinessCaseStatus,
                                        ovrlsts: result.OverallStatus,
                                        scoptrnd: result.ScopeTrend,
                                        schdltrnd: result.ScheduleTrend,
                                        bstrnd: result.BusinessCaseTrend,
                                        ovrltrnd: result.OveralTrend,
                                        cncsts: result.ChangecommasStatus,
                                        cnctrnd: result.ChangecommasTrend,
                                        imopsts: result.ImpactonOperationsStatus,
                                        imoptrnd: result.ImpactonOperationsTrend,
                                        achievements: result.keyachievementsinperiod,
                                        activities: result.keyactivitiesfornextperiod,
                                        supportnd: result.supportattentionneeded,
                                        id: result.ID,
                                        key: result.ID,
                                        name: result.ID
                                    });
                                    //console.log(items.Title);
                                });

                            });
                            resolve(items);
                        }, (error: any): void => {
                            reject("error ocuured" + error);
                        }
                    );
                }
            );
        });
    }

    public getlistitems(context: WebPartContext): Promise<ISubmission[]> {
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/Items?";
        geturl = geturl + "&select=Programs,Initiative,ScopeStatus,ScheduleStatus,BusinessCaseStatus,OverallStatus,ScopeTrend,ID,Author,ScheduleTrend,BusinessCaseTrend,OveralTrend,ChangecommasStatus,ChangecommasTrend,ImpactonOperationsStatus,ImpactonOperationsTrend,keyachievementsinperiod,keyactivitiesfornextperiod,supportattentionneeded,Modified";
        geturl = geturl + "&$orderby=Initiative";
        var items: ISubmission[] = [];
        return new Promise<ISubmission[]>(async (resolve, reject) => {
            context.spHttpClient
                .get(geturl, SPHttpClient.configurations.v1).then(
                    (Response: SPHttpClientResponse) => {
                        Response.json().then((results: any) => {
                            console.log('list arry', results);
                            results.value.map((result: any, index) => {
                                items.push({
                                    program: result.Programs,
                                    intiate: result.Initiative,
                                    scopsts: result.ScopeStatus,
                                    schdlsts: result.ScheduleStatus,
                                    bssts: result.BusinessCaseStatus,
                                    ovrlsts: result.OverallStatus,
                                    scoptrnd: result.ScopeTrend,
                                    schdltrnd: result.ScheduleTrend,
                                    bstrnd: result.BusinessCaseTrend,
                                    ovrltrnd: result.OveralTrend,
                                    cncsts: result.ChangecommasStatus,
                                    cnctrnd: result.ChangecommasTrend,
                                    imopsts: result.ImpactonOperationsStatus,
                                    imoptrnd: result.ImpactonOperationsTrend,
                                    achievements: result.keyachievementsinperiod,
                                    activities: result.keyactivitiesfornextperiod,
                                    supportnd: result.supportattentionneeded,
                                    id: result.ID,
                                    key: "item" + index,
                                    name: result.ID,
                                    auther: result.AuthorId,
                                    modified: result.Modified


                                });
                                //console.log(items.Title);
                            });

                        });
                        resolve(items);
                    }, (error: any): void => {
                        reject("error ocuured" + error);
                    }
                );
        });
    }

    public deleteItem(context: WebPartContext, selecteditems: IObjectWithKey[]): Promise<ISubmission[]> {
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/items";
        var items: ISubmission[] = [];
        var close: boolean = false;
        const options: ISPHttpClientOptions = {
            headers: {
                'Accept': "application/json;odata=nometadata",
                "content-type": "application/json;odata=nometadat",
                "odataverion": "",
                'IF-MATCH': "*",
                'X-HTTP-Method': 'DELETE',
            },
        };
        return new Promise<ISubmission[]>(async (resolve, reject) => {
            selecteditems.map((selecteditem: any) => {
                context.spHttpClient.post(geturl + "(" + selecteditem.id + ")", SPHttpClient.configurations.v1, options).then(
                    () => {
                        context.spHttpClient.get(geturl, SPHttpClient.configurations.v1).then(
                            (Response: SPHttpClientResponse) => {
                                Response.json().then((results: any) => {
                                    results.value.map((result: any) => {
                                        items.push({
                                            program: result.Programs,
                                            intiate: result.Initiative,
                                            scopsts: result.ScopeStatus,
                                            schdlsts: result.ScheduleStatus,
                                            bssts: result.BusinessCaseStatus,
                                            ovrlsts: result.OverallStatus,
                                            scoptrnd: result.ScopeTrend,
                                            schdltrnd: result.ScheduleTrend,
                                            bstrnd: result.BusinessCaseTrend,
                                            ovrltrnd: result.OveralTrend,
                                            cncsts: result.ChangecommasStatus,
                                            cnctrnd: result.ChangecommasTrend,
                                            imopsts: result.ImpactonOperationsStatus,
                                            imoptrnd: result.ImpactonOperationsTrend,
                                            achievements: result.keyachievementsinperiod,
                                            activities: result.keyactivitiesfornextperiod,
                                            supportnd: result.supportattentionneeded,
                                            id: result.ID,
                                            key: result.ID,
                                            name: result.ID,



                                        });
                                        //console.log(items.Title);
                                    });

                                });
                                resolve(items);
                            }, (error: any): void => {
                                reject("error ocuured" + error);
                            }
                        );
                    }
                );
            });
        });
    }

    public updateItem(context: WebPartContext, _listinitiate: ISubmission, selecteditems: IObjectWithKey[]): Promise<ISubmission[]> {
        let posturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/items";

        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/Items?";
        geturl = geturl + "&select=*";
        geturl = geturl + "&$orderby=Initiative";

        var items: ISubmission[] = [];
        var close: boolean = false;
        const body: string = JSON.stringify({
            Programs: _listinitiate.program,
            Initiative: _listinitiate.intiate,
            keyachievementsinperiod: _listinitiate.achievements,
            keyactivitiesfornextperiod: _listinitiate.activities,
            supportattentionneeded: _listinitiate.supportnd,
            ScopeStatus: _listinitiate.scopsts,
            ScheduleStatus: _listinitiate.schdlsts,
            BusinessCaseStatus: _listinitiate.bssts,
            OverallStatus: _listinitiate.ovrlsts,
            ScopeTrend: _listinitiate.scoptrnd,
            ScheduleTrend: _listinitiate.schdltrnd,
            BusinessCaseTrend: _listinitiate.bstrnd,
            OveralTrend: _listinitiate.ovrltrnd,
            ChangecommasStatus: _listinitiate.cncsts,
            ChangecommasTrend: _listinitiate.cnctrnd,
            ImpactonOperationsStatus: _listinitiate.imopsts,
            ImpactonOperationsTrend: _listinitiate.imoptrnd,

        });
        const options: ISPHttpClientOptions = {
            headers: {
                'Accept': "application/json;odata=nometadata",
                "content-type": "application/json;odata=nometadat",
                "odataverion": "",
                'IF-MATCH': "*",
                'X-HTTP-Method': 'MERGE',
            },
            body: body
        };
        return new Promise<ISubmission[]>(async (resolve, reject) => {
            selecteditems.map((selecteditem: any) => {
                context.spHttpClient.post(posturl + "(" + selecteditem.id + ")", SPHttpClient.configurations.v1, options).then(
                    () => {
                        context.spHttpClient.get(geturl, SPHttpClient.configurations.v1).then(
                            (Response: SPHttpClientResponse) => {
                                Response.json().then((results: any) => {
                                    results.value.map((result: any) => {
                                        items.push({
                                            program: result.Programs,
                                            intiate: result.Initiative,
                                            scopsts: result.ScopeStatus,
                                            schdlsts: result.ScheduleStatus,
                                            bssts: result.BusinessCaseStatus,
                                            ovrlsts: result.OverallStatus,
                                            scoptrnd: result.ScopeTrend,
                                            schdltrnd: result.ScheduleTrend,
                                            bstrnd: result.BusinessCaseTrend,
                                            ovrltrnd: result.OveralTrend,
                                            cncsts: result.ChangecommasStatus,
                                            cnctrnd: result.ChangecommasTrend,
                                            imopsts: result.ImpactonOperationsStatus,
                                            imoptrnd: result.ImpactonOperationsTrend,
                                            achievements: result.keyachievementsinperiod,
                                            activities: result.keyactivitiesfornextperiod,
                                            supportnd: result.supportattentionneeded,
                                            id: result.ID,
                                            key: result.ID,
                                            name: result.ID
                                        });
                                        //console.log(items.Title);
                                    });

                                });
                                resolve(items);
                            }, (error: any): void => {
                                reject("error ocuured" + error);
                            }
                        );
                    }
                );
            });
        });
    }

    public getinitiativeitems(context: WebPartContext, program: any): Promise<Analysis> {

        //let geturl: string = context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('Datacollect')/Items?";
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/items?";
        geturl = geturl + "&select=*";
        geturl = geturl + "&$filter=Programs eq '" + program + "'";
        geturl = geturl + "&$orderby=Initiative";
        var data: Array<number> = []
        var items: ISubmission[] = []; 0
        let analysis = new Analysis();
        return new Promise<Analysis>(async (resolve, reject) => {
            context.spHttpClient
                .get(geturl, SPHttpClient.configurations.v1).then(
                    (Response: SPHttpClientResponse) => {
                        Response.json().then((results: any) => {
                            analysis.Initiative = results.value;
                            results.value.map((result: any, index) => {
                                // scope
                                switch (result.ScopeStatus) {
                                    case 'minor issues threatening scheduleand/or goals': {
                                        analysis.Data[0].Status.datasets[0].data[0] += 1;
                                        break;
                                    }
                                    case 'On schedule;goals within reach': {
                                        analysis.Data[0].Status.datasets[0].data[1] += 1;
                                        break;
                                    }
                                    case 'Behind schedule and/or goals are risk': {
                                        analysis.Data[0].Status.datasets[0].data[2] += 1;
                                        break;
                                    }
                                }
                                switch (result.ScopeTrend) {
                                    case 'Stable': {
                                        analysis.Data[0].Trend.datasets[0].data[0] += 1;
                                        break;
                                    }
                                    case 'Trending up': {
                                        analysis.Data[0].Trend.datasets[0].data[1] += 1;
                                        break;
                                    }
                                    case 'Trending down': {
                                        analysis.Data[0].Trend.datasets[0].data[2] += 1;
                                        break;
                                    }
                                }
                                // schedule
                                switch (result.ScheduleStatus) {
                                    case 'minor issues threatening scheduleand/or goals': {
                                        analysis.Data[1].Status.datasets[0].data[0] += 1;
                                        break;
                                    }
                                    case 'On schedule;goals within reach': {
                                        analysis.Data[1].Status.datasets[0].data[1] += 1;
                                        break;
                                    }
                                    case 'Behind schedule and/or goals are risk': {
                                        analysis.Data[1].Status.datasets[0].data[2] += 1;
                                        break;
                                    }
                                }
                                switch (result.ScheduleTrend) {
                                    case 'Stable': {
                                        analysis.Data[1].Trend.datasets[0].data[0] += 1;
                                        break;
                                    }
                                    case 'Trending up': {
                                        analysis.Data[1].Trend.datasets[0].data[1] += 1;
                                        break;
                                    }
                                    case 'Trending down': {
                                        analysis.Data[1].Trend.datasets[0].data[2] += 1;
                                        break;
                                    }
                                }
                                // business
                                switch (result.BusinessCaseStatus) {
                                    case 'minor issues threatening scheduleand/or goals': {
                                        analysis.Data[2].Status.datasets[0].data[0] += 1;
                                        break;
                                    }
                                    case 'On schedule;goals within reach': {
                                        analysis.Data[2].Status.datasets[0].data[1] += 1;
                                        break;
                                    }
                                    case 'Behind schedule and/or goals are risk': {
                                        analysis.Data[2].Status.datasets[0].data[2] += 1;
                                        break;
                                    }
                                }
                                switch (result.BusinessCaseTrend) {
                                    case 'Stable': {
                                        analysis.Data[2].Trend.datasets[0].data[0] += 1;
                                        break;
                                    }
                                    case 'Trending up': {
                                        analysis.Data[2].Trend.datasets[0].data[1] += 1;
                                        break;
                                    }
                                    case 'Trending down': {
                                        analysis.Data[2].Trend.datasets[0].data[2] += 1;
                                        break;
                                    }
                                }
                                // overall
                                switch (result.OverallStatus) {
                                    case 'minor issues threatening scheduleand/or goals': {
                                        analysis.Data[3].Status.datasets[0].data[0] += 1;
                                        break;
                                    }
                                    case 'On schedule;goals within reach': {
                                        analysis.Data[3].Status.datasets[0].data[1] += 1;
                                        break;
                                    }
                                    case 'Behind schedule and/or goals are risk': {
                                        analysis.Data[3].Status.datasets[0].data[2] += 1;
                                        break;
                                    }
                                }
                                switch (result.OveralTrend) {
                                    case 'Stable': {
                                        analysis.Data[3].Trend.datasets[0].data[0] += 1;
                                        break;
                                    }
                                    case 'Trending up': {
                                        analysis.Data[3].Trend.datasets[0].data[1] += 1;
                                        break;
                                    }
                                    case 'Trending down': {
                                        analysis.Data[3].Trend.datasets[0].data[2] += 1;
                                        break;
                                    }
                                }

                            });
                        });
                        resolve(analysis);
                    }, (error: any): void => {
                        reject("error ocuured" + error);
                    }
                );
        });
    }
    public getRecentInitiative(context: WebPartContext, program: IDropdownOption[]): Promise<any> {
        var items: any = [];
        return new Promise<Analysis>(async (resolve, reject) => {
            await program.map(async (value: IDropdownOption) => {
                let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/items?";
                geturl = geturl + "&select=*";
                geturl = geturl + "&$filter=Programs eq '" + value.text + "'";
                geturl = geturl + "&$orderby=Created&$top=1";
                await context.spHttpClient
                    .get(geturl, SPHttpClient.configurations.v1).then(
                        (Response: SPHttpClientResponse) => {
                            Response.json().then((results: any) => {
                                console.log('recent records for need attention', results);
                                if (results.value[0])
                                    items.push(results.value[0]);
                            });
                        }
                    );
            })
            resolve(items);
        })
    }
    public needAttentionList(context: WebPartContext): Promise<any> {
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/Items?";
        geturl = geturl + "&select=Programs,Initiative,ScopeStatus,ScheduleStatus,BusinessCaseStatus,OverallStatus,ScopeTrend,ID,Author,ScheduleTrend,BusinessCaseTrend,OveralTrend,ChangecommasStatus,ChangecommasTrend,ImpactonOperationsStatus,ImpactonOperationsTrend,keyachievementsinperiod,keyactivitiesfornextperiod,supportattentionneeded,Modified";
        geturl = geturl + "&$orderby=Modified desc&$top=10";
        var items: any = [];
        let count = 0;
        return new Promise<any>(async (resolve, reject) => {
            context.spHttpClient
                .get(geturl, SPHttpClient.configurations.v1).then(
                    (Response: SPHttpClientResponse) => {
                        Response.json().then((results: any) => {
                            results.value.map(x => {
                                if (x.supportattentionneeded != '' && count < 5) {
                                    items.push(x);
                                    count++;
                                }
                            })
                        });
                        resolve(items);
                    }, (error: any): void => {
                        reject("error ocuured" + error);
                    }
                );
        });
    }

    public getReports(context: WebPartContext, program: any): Promise<any> {
        let reportList = new Array<Category>();
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/items?";
        geturl = geturl + "&select=*";
        geturl = geturl + "&$filter=Programs eq '" + program + "'";
        geturl = geturl + "&$orderby=Initiative";
        return new Promise<any>(async (resolve, reject) => {
            context.spHttpClient
                .get(geturl, SPHttpClient.configurations.v1).then(
                    (Response: SPHttpClientResponse) => {
                        Response.json().then((results: any) => {
                            results.value.map((x: any, index) => {
                                const report = new Category();
                                if (index == 0) {
                                    report.Status = true;
                                    report.Count = results.value.length;
                                }
                                report.Programs = x.Programs;
                                report.Initiative = x.Initiative;
                                report.ScopeTrend = x.ScopeTrend;
                                report.ScheduleTrend = x.ScheduleTrend;
                                report.BusinessCaseTrend = x.BusinessCaseTrend;
                                report.OveralTrend = x.OveralTrend;
                                report.ScopeStatus = x.ScopeStatus;
                                report.ScheduleStatus = x.ScheduleStatus;
                                report.BusinessCaseStatus = x.BusinessCaseStatus;
                                report.OverallStatus = x.OverallStatus;
                                report.Created = moment(x.Created, 'DD-MM-YYYY').format('YYYY-MM-DD').toString();;
                                reportList.push(report);
                            })
                            resolve(reportList);
                        });
                    }, (error: any): void => {
                        reject("error ocuured" + error);
                    }
                );
        });
    }
    public getAllReports(context: WebPartContext, program: IDropdownOption[]): Promise<any> {
        console.log('service test');
        let allReports: any = [];
        let reportList = new Array<Category>();

        return new Promise<any>(async (resolve, reject) => {
            await program.map(async (value: IDropdownOption) => {
                let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/items?";
                geturl = geturl + "&select=*";
                geturl = geturl + "&$filter=Programs eq '" + value.text + "'";
                geturl = geturl + "&$orderby=Initiative";
                await context.spHttpClient
                    .get(geturl, SPHttpClient.configurations.v1).then(
                        (Response: SPHttpClientResponse) => {
                            Response.json().then((results: any) => {
                                results.value.map((z: any, index) => {
                                    const report = new Category();
                                    if (index == 0) {
                                        report.Status = true;
                                        report.Count = results.value.length;
                                    }
                                    report.Programs = z.Programs;
                                    report.Initiative = z.Initiative;
                                    report.ScopeTrend = z.ScopeTrend;
                                    report.ScheduleTrend = z.ScheduleTrend;
                                    report.BusinessCaseTrend = z.BusinessCaseTrend;
                                    report.OveralTrend = z.OveralTrend;
                                    report.ScopeStatus = z.ScopeStatus;
                                    report.ScheduleStatus = z.ScheduleStatus;
                                    report.BusinessCaseStatus = z.BusinessCaseStatus;
                                    report.OverallStatus = z.OverallStatus;
                                    report.Created = moment(z.Created, 'DD-MM-YYYY').format('YYYY-MM-DD').toString();;;
                                    reportList.push(report);
                                })
                            });

                        }, (error: any): void => {
                            reject("error ocuured" + error);
                        }
                    );
            });
            resolve(reportList);
        });
    }
    public getTrends(context: WebPartContext, initiative: any, count:number): Promise<any> {
        let reportList = new Array<Category>();
        let trendList = new TrendAnalysis();
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/lists/GetByTitle('Datacollect')/items?";
        geturl = geturl + "&select=*";
        geturl = geturl + "&$filter=Initiative eq '" + initiative + "'";
        geturl = geturl + "&$orderby=Created&$top="+count;
        return new Promise<any>(async (resolve, reject) => {
            context.spHttpClient
                .get(geturl, SPHttpClient.configurations.v1).then(
                    (Response: SPHttpClientResponse) => {
                        Response.json().then((results: any) => {
                            trendList.Initiative = results.value;
                            results.value.map((x: any, index) => {
                                console.log('trendddd', x);
                               trendList.Dates[index] = moment(x.Created, 'DD-MM-YYYY').format('MMMM Do').toString();
                                switch(x.ScopeTrend){
                                    case 'Trending up':{ 
                                        console.log('trening up0',trendList.Counts, index);
                                        trendList.Counts[index] =  trendList.Counts[index] + 1;break;
                                    }
                                    case 'Trending down': {trendList.Counts[index] =  trendList.Counts[index] + 1;break;}
                                    case 'Stable': {trendList.Counts[index] =  trendList.Counts[index] + 1;break;}
                                }
                            })
                            resolve(trendList);
                        });
                    }, (error: any): void => {
                        reject("error ocuured" + error);
                    }
                );
        });
    }
    public getCurrentUser(context: WebPartContext): Promise<number> {
        let user: any;
        const payload: string = JSON.stringify({
            'logonName': context.pageContext.user.email
        });
        var postData: ISPHttpClientOptions = {
            body: payload
        };
        let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/web/ensureuser";
        return new Promise<number>(async (resolve, reject) => {
            context.spHttpClient
                .post(geturl, SPHttpClient.configurations.v1, postData).then(
                    (Response: SPHttpClientResponse) => {
                        Response.json().then((results: any) => {
                            console.log('user profile', results);
                            resolve(results.Id);
                        });
                    }, (error: any): void => {
                        reject("error ocuured" + error);
                    }
                );
        });
    }


    test() {
        // get shartrepoint logged in user

        // let geturl: string = "https://sticsoftio.sharepoint.com/sites/poc/_api/SP.UserProfiles.PeopleManager/GetMyProperties";
        // return new Promise<number>(async (resolve, reject) => {
        //     context.spHttpClient
        //         .get(geturl, SPHttpClient.configurations.v1).then(
        //             (Response: SPHttpClientResponse) => {
        //                 Response.json().then((results: any) => {
        //                     console.log('user profile',results);
        //                     resolve(results.Id);
        //                 });
        //             }, (error: any): void => {
        //                 reject("error ocuured" + error);
        //             }
        //         );
        // });
    }

}