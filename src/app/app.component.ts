import { Component, ViewChild, ElementRef } from '@angular/core';
import * as XLSX from 'xlsx';
import jspdf from 'jspdf';

import html2canvas from 'html2canvas';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  title = 'angular-app';
  today = new Date();
  weekdays = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  monthNames = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ];
  day = this.weekdays[this.today.getDay()];
  month = this.monthNames[this.today.getMonth()];

  time =
    this.today.getDate() +
    '_' +
    this.month +
    '_' +
    this.today.getHours() +
    ':' +
    this.today.getMinutes();

  time2 = Math.floor(this.today.getTime() / 1000);

  fileName = this.time2 + '.xlsx';
  pdfName = this.time2 + '.pdf';
  userList = [
    {
      id: 1,
      name: 'lorem ipsum',
      username: 'lorem',
      email: 'Lorem@april.biz',
    },
    {
      id: 2,
      name: 'Shoaib',
      username: 'xoaeb',
      email: 'shoaib100@gmail.tv',
    },
    {
      id: 3,
      name: 'Clementine Bauch',
      username: 'Samantha',
      email: 'Nathan@yesenia.net',
    },
    {
      id: 4,
      name: 'Patricia Lebsack',
      username: 'Karianne',
      email: 'Julianne.OConner@kory.org',
    },
    {
      id: 5,
      name: 'Roshan Thakur',
      username: 'rossshan',
      email: 'roshanthakur@gmail.com',
    },
  ];

  exportexcel(): void {
    let element = document.getElementById('excel-table');
    const ws: XLSX.WorkSheet = XLSX.utils.table_to_sheet(element);

    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    XLSX.writeFile(wb, this.fileName);
  }

  @ViewChild('content') content!: ElementRef;

  public openPDF(): void {
    let DATA: any = document.getElementById('content');
    html2canvas(DATA).then((canvas) => {
      let fileWidth = 200;
      let fileHeight = 150;
      const FILEURI = canvas.toDataURL('image/png');
      let PDF = new jspdf('p', 'mm', 'a4');
      let position = 0;
      PDF.addImage(FILEURI, 'PNG', 0, position, fileWidth, fileHeight);
      PDF.save(this.pdfName);
    });
  }
}
