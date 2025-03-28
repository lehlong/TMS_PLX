import {Component, Pipe} from '@angular/core';
import {ActionLogService} from '../../services/system-manager/action-log.service';
import {GlobalService} from '../../services/global.service';
import {PaginationResult} from '../../models/base.model';
import {ShareModule} from '../../shared/share-module';
import {ActionLogFilter} from '../../models/master-data/action-log.model';
import {getISOWeek} from 'date-fns';
import {startOfWeek, endOfWeek, addDays, startOfDay, endOfDay, addHours, format} from 'date-fns';
import {PrettyPrintPipe} from '../../shared/custom-pipe/PrettyPrint/pretty-print.pipe';

@Component({
  selector: 'app-action-log',
  standalone: true,
  imports: [ShareModule, PrettyPrintPipe],
  templateUrl: './action-log.component.html',
  styleUrl: './action-log.component.scss',
})
export class ActionLogComponent {
  filter = new ActionLogFilter();
  paginationResult = new PaginationResult();
  rangePresets = {
    '2 giờ sau': [new Date(), addHours(new Date(), 2)],
    'Hôm nay': [startOfDay(new Date()), endOfDay(new Date())],
    'Ngày mai': [startOfDay(addDays(new Date(), 1)), endOfDay(addDays(new Date(), 1))],
    '3 ngày sau': [new Date(), endOfDay(addDays(new Date(), 3))],
    'Tuần này': [startOfWeek(new Date()), endOfWeek(new Date())],
  };
  edit: boolean = false;
  visible: boolean = false;
  dataDetail: any = null;
  listStatusCode = [
    {code: 200, message: 'Success'},
    {code: 400, message: 'Bad request'},
    {code: 500, message: 'Internal Server Error'},
  ];

  constructor(private _service: ActionLogService, private globalService: GlobalService) {
    this.globalService.setBreadcrumb([
      {
        name: 'Lịch sử thao tác',
        path: 'system-manager/action-log',
      },
    ]);
  }

  ngOnDestroy() {
    this.globalService.setBreadcrumb([]);
  }

  ngOnInit(): void {
    this.search();
  }

  search() {
    let filterFormat = {
      currentPage: this.filter.currentPage,
      pageSize: this.filter.pageSize,
      keyWord: this.filter.keyWord,
      FromDate:
        this.filter?.selectedRange?.length > 0 ? format(this.filter?.selectedRange[0], "yyyy-MM-dd'T'HH:mm:ss") : '',
      ToDate:
        this.filter?.selectedRange?.length > 1 ? format(this.filter?.selectedRange[1], "yyyy-MM-dd'T'HH:mm:ss") : '',
      StatusCode: this.filter.statusCode,
    };
    this._service.Search(filterFormat).subscribe({
      next: (data) => {
        this.paginationResult = data;
      },
      error: (response) => {
        console.log(response);
      },
    });
  }

  reset() {
    this.filter = new ActionLogFilter();
    this.search();
  }

  pageSizeChange(size: number): void {
    this.filter.currentPage = 1;
    this.filter.pageSize = size;
    this.search();
  }

  pageIndexChange(index: number): void {
    this.filter.currentPage = index;
    this.search();
  }

  onChange(result: Date[]): void {
    console.log('onChange: ', result);
  }

  getWeek(result: Date[]): void {
    console.log('week: ', result.map(getISOWeek));
  }

  openEdit(data: any) {
    this.visible = true;
    this.dataDetail = data;
  }

  close() {
    this.visible = false;
  }
}
