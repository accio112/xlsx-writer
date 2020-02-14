import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { XlsxWriterDirective } from './xlsx-writer.directive';
import { DatePipe } from '@angular/common';
@NgModule({
  declarations: [
    AppComponent,
    XlsxWriterDirective
  ],
  imports: [
    BrowserModule,
    AppRoutingModule
  ],
  providers: [DatePipe],
  bootstrap: [AppComponent]
})
export class AppModule { }
