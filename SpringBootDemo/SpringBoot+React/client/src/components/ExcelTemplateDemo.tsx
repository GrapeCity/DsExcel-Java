import * as React from 'react';
import {RouteComponentProps} from "react-router";
import 'isomorphic-fetch';
import { Utility } from '../utility'


//ExcelTemplate
export class ExcelTemplateDemo extends React.Component<{} & RouteComponentProps<{}>, {}> {
    templateName: string;
    spread: GC.Spread.Sheets.Workbook;

    constructor() {
        super(null);

        // this is a template excel file at server side
        this.templateName = 'SimpleInvoice.xlsx';

        this.exportExcel = this.exportExcel.bind(this);
    }

    public render() {
        return <div className='spread-page'>
            <h1>Excel Template Demo</h1>
            <p>This example demonstrates how to use <strong>GcExcel</strong> as server spreadsheet model, and use <strong>Spread.Sheets</strong> as client side viewer or editor: </p>
            <ul>
                <li><strong>GcExcel</strong> will first open an excel file existing on server.</li>
                <li><strong>GcExcel</strong> then inoke <strong>ToJson</strong> and transport the ssjson to client side.</li>
                <li>In browser, <strong>Spread.Sheets</strong> will invoke <strong>fromJSON</strong> with the ssjson from server.</li>
                <li>Then, you can see the content of the excel template in <strong>Spread.Sheets</strong>.</li>
                <li>At same time, you can fill or edit data on the template through <strong>Spread.Sheets</strong>.</li>
                <li>At last, you can download to view the changes in excel file.</li>
            </ul>
            <br />
            <div className='btn-group'>
                <button className='btn btn-default btn-md' onClick={this.exportExcel}>Export Excel</button>
            </div>
            <div id='spreadjs' className='spread-div' />
        </div>;
    }

    componentDidMount() {
        this.spread = new GC.Spread.Sheets.Workbook(document.getElementById('spreadjs'), {
            seetCount: 1
        });

        this.loadSpreadFromTemplate();
    }

    loadSpreadFromTemplate() {
        var requestUrl = '/template/' + this.templateName;
        fetch(requestUrl, {
            method: 'Get'
        }).then(response => response.json() as Promise<object>)
            .then(data => {
                this.spread.fromJSON(data);
            });
    }

    exportExcel() {
        var ssjson = JSON.stringify(this.spread.toJSON(null));
        Utility.ExportExcel(ssjson, this.templateName);
    }

}


