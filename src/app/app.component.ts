import { Component } from '@angular/core';
import {createDataOne} from './test-data/test-data1';
import {createDataTwo} from './test-data/test-data2';
import {createDataThree} from './test-data/test-data3';
import {createDataFour} from './test-data/test-data4';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent{
 
  title = 'xlsx-writer';
  public enableDownload1 = false;
  public enableDownload2 = false;
  public enableDownload3 = false;
  public enableDownload4 = false;
  constructor() {
  }
  downloadData1 = createDataOne();
  downloadData2 = createDataTwo();
  downloadData3 = createDataThree();
  downloadData4 = createDataFour();

  download1(){
    this.enableDownload1 = true;
  }
  download2(){
    this.enableDownload2 = true;
  }
  download3(){
    this.enableDownload3 = true;
  }
  download4(){
    this.enableDownload4 = true;
  }

}
