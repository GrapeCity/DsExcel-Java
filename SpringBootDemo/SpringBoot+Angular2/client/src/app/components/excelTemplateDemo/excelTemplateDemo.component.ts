import { Component, Inject } from '@angular/core';
import { Http } from '@angular/http';
import { Utility } from '../../utility';

@Component({
    selector: 'excelTemplateDemo',
    templateUrl: './excelTemplateDemo.component.html'
})

export class ExcelTemplateDemoComponent {
    spread: GC.Spread.Sheets.Workbook;
    templateName: string;

    constructor(http: Http, @Inject('ORIGIN_URL') originUrl: string) {

        // this is a template excel file at server side
        this.templateName = 'SimpleInvoice.xlsx';
    }

    getSpread($event) {
        this.spread = $event.spread;
        this.loadSpreadFromTemplate();
    }

    private loadSpreadFromTemplate() : void {
        var requestUrl = '/template/' + this.templateName;
        fetch(requestUrl, {
            method: 'Get'
        }).then(response => response.json() as Promise<object>)
            .then(data => {
                this.spread.fromJSON(data);
            });
    }

    public exportExcel(): void {
        var ssjson = JSON.stringify(this.spread.toJSON(null));
        Utility.ExportExcel(ssjson, this.templateName);
    }
}
