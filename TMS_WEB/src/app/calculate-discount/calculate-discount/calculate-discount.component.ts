import { Component, OnInit } from '@angular/core';
import { ShareModule } from '../../shared/share-module';
import { CalculateDiscountService } from '../../services/calculate-discount/calculate-discount.service';
import { GlobalService } from '../../services/global.service';
import { NzMessageService } from 'ng-zorro-antd/message';
import { BaseFilter, PaginationResult } from '../../models/base.model';
import { Router } from '@angular/router';

@Component({
  selector: 'app-calculate-discount',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './calculate-discount.component.html',
  styleUrl: './calculate-discount.component.scss'
})
export class CalculateDiscountComponent implements OnInit {
  constructor(
    private _service: CalculateDiscountService,
    private globalService: GlobalService,
    private message: NzMessageService,
    private router: Router,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách đợt tính mức giảm giá',
        path: 'calculate-discount/list',
      },
    ])
    this.globalService.getLoading().subscribe((value) => {
      this.loading = value
    })
  }
  isVisibleStatus: boolean = false
  noData: any[] = []
  loading: boolean = false
  visible = false;
  filter = new BaseFilter()
  paginationResult = new PaginationResult()

  statusModel = {
    title: '',
    des: '',
    value: '',
  }
  input: any = {
    header: {},
    inputPrice: [],
    market: [],
    customerDb: [],
    customerPt: [],
    customerFob: [],
    customerTnpp: [],
    customerBbdo: [],
  };

  ngOnInit(): void {
    this.search();
  
  }
  search() {
    this._service.search(this.filter).subscribe({
      next: (data) => {
        this.paginationResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  changeStatus(value: string, status: string) {
    switch (value) {
      case '01':
        this.statusModel.title = 'TRÌNH DUYỆT'
        this.statusModel.des = 'Bạn có muốn Trình duyệt dữ liệu này?'
        break
      case '02':
        this.statusModel.title = 'YÊU CẦU CHỈNH SỬA'
        this.statusModel.des = 'Bạn có muốn Yêu cầu chỉnh sửa lại dữ liệu này?'
        break
      case '03':
        this.statusModel.title = 'PHÊ DUYỆT'
        this.statusModel.des = 'Bạn có muốn Phê duyệt dữ liệu này?'
        break
      case '04':
        this.statusModel.title = 'TỪ CHỐI'
        this.statusModel.des = 'Bạn có muốn Từ chối dữ liệu này?'
        break
      case '05':
        this.statusModel.title = 'HỦY TRÌNH DUYỆT'
        this.statusModel.des = 'Bạn có muốn Hủy trình duyệt dữ liệu này?'
        break
      case '06':
        this.statusModel.title = 'HỦY PHÊ DUYỆT'
        this.statusModel.des = 'Bạn có muốn Hủy phê duyệt dữ liệu này?'
        break
    }
    this.input.status.code = status
    this.isVisibleStatus = true
  }
  genarateCreate() {
    this._service.genarateCreate().subscribe({
      next: (data) => {
        this.input = data;
        this.visible = true;
        console.log(data)
      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  onCreate() {
    this._service.create(this.input).subscribe({
      next: (data) => {
        this.router.navigate([`/calculate-discount/detail/${this.input.header.id}`]);
      },
      error: (response) => {
        console.log(response)
      },
    });
  }
  openCalculateDetail(id: any) {
    this.router.navigate([`/calculate-discount/detail/${id}`]);
  }
  close(): void {
    this.visible = false;
    this.input = {
      header: {},
      inputPrice: [],
      market: [],
      customerDb: [],
      customerPt: [],
      customerFob: [],
      customerTnpp: [],
    };
  }
  onChange(result: Date): void {
    console.log('Selected Time: ', result);
  }

  onOk(result: Date | Date[] | null): void {
    console.log('onOk', result);
  }
}
