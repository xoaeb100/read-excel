import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

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
    ' ' +
    this.month +
    ' ' +
    this.today.getHours() +
    ':' +
    this.today.getMinutes();
  fileName = this.time + '.xlsx';
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
}
