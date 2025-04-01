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
