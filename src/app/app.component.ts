import {Component} from '@angular/core';
import {HttpClient} from '@angular/common/http';
import "@grapecity/spread-sheets-resources-zh";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'ng-app';
  spreadBackColor = 'aliceblue';
  sheetName = 'Goods List';
  spread;
  json;
  excelIo;
  hostStyle = {
    width: '100%',
    height: '600px'
  };
  data = [
    {Name: 'Apple', Category: 'Fruit', Price: 1, 'Shopping Place': 'Wal-Mart'},
    {Name: 'Potato', Category: 'Fruit', Price: 2.01, 'Shopping Place': 'Other'},
    {Name: 'Tomato', Category: 'Vegetable', Price: 3.21, 'Shopping Place': 'Other'},
    {Name: 'Sandwich', Category: 'Food', Price: 2, 'Shopping Place': 'Wal-Mart'},
    {Name: 'Hamburger', Category: 'Food', Price: 2, 'Shopping Place': 'Wal-Mart'},
    {Name: 'Grape', Category: 'Fruit', Price: 4, 'Shopping Place': 'Sun Store'}
  ];
  columnWidth = 100;

  constructor(private http: HttpClient) {
    this.http.get('/assets/xlxs.json').subscribe(v => {
      console.log(v);
      this.json = v;
      this.spread.fromJSON(this.json);
    });
  }

  workbookInit(args) {
    let options = {
      backColor: 'red'
    };
    this.spread = args.spread;
    this.spread.options.grayAreaBackColor = 'white';
    this.excelIo = new GC.Spread.Excel.IO();
    GC.Spread.Common.CultureManager.culture("zh-cn");
  }

  upload(e) {
    let file = e.target.files[0];
    let that = this;
    this.excelIo.open(file, function(json) {
      var workbookObj = json;
      console.log(JSON.stringify(json));
      that.spread.fromJSON(workbookObj);
    }, function(e) {
    }, {password: 123});
  }

  export() {

    var json = this.spread.toJSON();
    console.log(JSON.stringify(json));
    // here is excel IO API
    let that = this;
    this.excelIo.save(json, function(blob) {
      that.downloadContent(blob, '123.xlsx');
    }, function(e) {
      // process error
      console.log(e);
    });

  }

  addSheet() {
    this.spread.addSheet(this.spread.getSheetCount());

  }

  downloadContent(content, fileName?: string) {
    let eleLink = document.createElement('a');
    eleLink.download = fileName || '下载.xlsx';
    eleLink.style.display = 'none';

    const blob = new Blob([content]);
    eleLink.href = URL.createObjectURL(blob);

    document.body.appendChild(eleLink);
    eleLink.click();
    document.body.removeChild(eleLink);
  }


}
