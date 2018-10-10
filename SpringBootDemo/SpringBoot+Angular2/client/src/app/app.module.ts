import { NgModule } from '@angular/core';
import { RouterModule } from '@angular/router';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import { HttpModule } from '@angular/http';

import { AppComponent } from './components/app/app.component';
import { NavMenuComponent } from './components/navmenu/navmenu.component';
import { HomeComponent } from './components/home/home.component';
import { ExcelTemplateDemoComponent } from './components/excelTemplateDemo/excelTemplateDemo.component';
import { ProgrammingDemoComponent } from './components/programmingDemo/programmingDemo.component';
import { ExcelIODemoComponent } from './components/excelIODemo/excelIODemo.component';
import { SpreadSheetsModule } from './gc.spread.sheets.angular2.10.2.0';

@NgModule({
  declarations: [
    AppComponent,
    NavMenuComponent,
    ExcelIODemoComponent,
    ExcelTemplateDemoComponent,
    ProgrammingDemoComponent,
    HomeComponent
  ],
  imports: [
    BrowserModule,
    FormsModule,
    HttpModule,
    SpreadSheetsModule,
    RouterModule.forRoot([
      { path: '', component: HomeComponent },
      { path: 'home', component: HomeComponent },
      { path: 'excelTemplateDemo', component: ExcelTemplateDemoComponent },
      { path: 'programmingDemo', component: ProgrammingDemoComponent },
      { path: 'excelIODemo', component: ExcelIODemoComponent },
      { path: '**', redirectTo: 'home' }
    ])
  ],
  providers: [{ provide: 'ORIGIN_URL', useValue: location.origin }],
  bootstrap: [AppComponent]
})
export class AppModule { }
