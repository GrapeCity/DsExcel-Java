import { Component, Inject } from '@angular/core';
import { Http } from '@angular/http';
import { Utility } from '../../utility';

@Component({
    selector: 'programmingDemo',
    templateUrl: './programmingDemo.component.html'
})

export class ProgrammingDemoComponent {
    spread: GC.Spread.Sheets.Workbook;
    options: Array<object> = [
        { value: 'BidTracker', label: 'Bid Tracker' },
        { value: 'ToDoList', label: 'ToDo List' },
        { value: 'AddressBook', label: 'Address Book' }
    ];
    selectedValue: string;

    constructor() {
        this.selectedValue = "BidTracker";
    }

    onChange(newValue) {
        console.log(newValue);
        this.selectedValue = newValue;
        this.loadSpreadFromUseCase(this.selectedValue);
    }

    getSpread($event) {
        this.spread = $event.spread;
        this.loadSpreadFromUseCase(this.selectedValue);
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

    public exportExcel(): void {
        var ssjson = JSON.stringify(this.spread.toJSON(null));
        Utility.ExportExcel(ssjson, this.selectedValue);
    }
}
