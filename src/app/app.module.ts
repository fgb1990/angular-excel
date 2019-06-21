import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';

import { A11yModule } from '@angular/cdk/a11y';
import { SpreadSheetsModule } from "@grapecity/spread-sheets-angular";
import {HttpClient, HttpClientModule} from '@angular/common/http';
@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    HttpClientModule,
    A11yModule,
    SpreadSheetsModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
