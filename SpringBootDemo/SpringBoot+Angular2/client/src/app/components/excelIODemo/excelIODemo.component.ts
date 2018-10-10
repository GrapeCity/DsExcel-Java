import { Component } from '@angular/core';
import { Utility } from '../../utility';

@Component({
    selector: 'excelIODemo',
    templateUrl: './excelIODemo.component.html',
    styleUrls: ['./excelIODemo.component.css']
})

export class ExcelIODemoComponent {
    spread: GC.Spread.Sheets.Workbook;
    selectedFileName: string;

    getSpread($event) {
        this.spread = $event.spread;

        var sheet = this.spread.getActiveSheet();
        sheet.addSpan(6, 3, 5, 10);
        var cell = sheet.getCell(6, 3);
        cell.text('Please choose an excel file(xlsx) to import!');
        cell.hAlign(GC.Spread.Sheets.HorizontalAlign.center);
        cell.vAlign(GC.Spread.Sheets.VerticalAlign.center);
        cell.font('bold 25px arial');
    }


    /**
     * Upload an excel file at client side, open the file at server side then transport the ssjson to client
     * @param e
     */
    public importExcel(e) : void {
        var selectedFile = e.target.files[0];
        if (!selectedFile) {
            this.selectedFileName = null;
            return;
        }

        this.selectedFileName = selectedFile.name;
        var requestUrl = '/open';
        fetch(requestUrl, {
            method: 'POST',
            body: selectedFile
        }).then(response => response.json() as Promise<object>)
            .then(data => {
                this.spread.fromJSON(data);
            });
    }

    /**
     * Tranport ssjson from Spread.Sheets and save and download the excel file.
     * @param e
     */
    public exportExcel(e) : void {
        var ssjson = JSON.stringify(this.spread.toJSON(null));
        Utility.ExportExcel(ssjson, this.selectedFileName);
    }
}
