import * as React from 'react';
import { RouteComponentProps } from "react-router";
import 'isomorphic-fetch';
import { Utility } from '../utility';
import Select from "react-select";
import '../css/react-select.css';

interface ProgrammingDemoState {
    options: { value: string; label: string; }[],
    value: { value: string; label: string; }
}

//ServerSidePrograming
export class ProgrammingDemo extends React.Component<{} & RouteComponentProps<{}>, ProgrammingDemoState> {
    spread: GC.Spread.Sheets.Workbook;

    constructor() {
        super(null);

        this.state = {
            options: [
                { value: 'BidTracker', label: 'Bid Tracker' },
                { value: 'ToDoList', label: 'ToDo List' },
                { value: 'AddressBook', label: 'Address Book' }
            ],
            value: { value: 'BidTracker', label: 'Bid Tracker' }
        }

        this.exportExcel = this.exportExcel.bind(this);
        this.onUseCaseChange = this.onUseCaseChange.bind(this);
    }

    public render() {
        return <div className='spread-page'>
            <h1>Programming API Demo</h1>
            <p>This example demonstrates how to programme with <strong>GcExcel</strong> to generate a complete spreadsheet model at server side, you can find all of source code in the SpreadServicesController.cs, we use <strong>Spread.Sheets</strong> as client side viewer. </p>            <ul>
                <li>You can first program with <strong>GcExcel</strong> at server side.</li>
                <li><strong>GcExcel</strong> then inoke <strong>ToJson</strong> and transport the ssjson to client side.</li>
                <li>In browser, <strong>Spread.Sheets</strong> will invoke <strong>fromJSON</strong> with the ssjson from server.</li>
                <li>Then, you can view the result in <strong>Spread.Sheets</strong> or download it as excel file.</li>
            </ul>
            <br />
            <div className='btn-group'>
                <Select className='select'
                    name="form-field-name"
                    value={this.state.value}
                    options={this.state.options}
                    onChange={this.onUseCaseChange} />
                <button className='btn btn-default btn-md' onClick={this.exportExcel}>Export Excel</button>
            </div>
            <div id='spreadjs' className='spread-div' />
        </div>;
    }

    componentDidMount() {
        this.spread = new GC.Spread.Sheets.Workbook(document.getElementById('spreadjs'), {
            seetCount: 1
        });

        this.loadSpreadFromUseCase(this.state.value.value);
    }

    loadSpreadFromUseCase(caseName: string) {
        var requestUrl = '/json/' + caseName;
        fetch(requestUrl, {
            method: 'Get'
        }).then(response => response.json() as Promise<object>)
            .then(data => {
                this.spread.fromJSON(data);
            });
    }

    onUseCaseChange(newValue: any) {
        this.setState({ value: newValue });
        this.loadSpreadFromUseCase(newValue.value);
    }

    exportExcel() {
        var ssjson = JSON.stringify(this.spread.toJSON(null));
        Utility.ExportExcel(ssjson, this.state.value.value);
    }
}


