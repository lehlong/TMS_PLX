import { Component, OnInit, output } from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { NzMessageService } from 'ng-zorro-antd/message';
import { CalculateDiscountService } from '../../services/calculate-discount/calculate-discount.service';
import { GlobalService } from '../../services/global.service';
import { ShareModule } from '../../shared/share-module';
import {
  CALCULATE_RESULT_RIGHT,
  IMPORT_BATCH,
} from '../../shared/constants/access-right.constants'
@Component({
  selector: 'app-calculate-discount-detail',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './calculate-discount-detail.component.html',
  styleUrl: './calculate-discount-detail.component.scss'
})
export class CalculateDiscountDetailComponent implements OnInit {
  titleTab: string = 'Dữ liệu gốc';
  loading: boolean = false;
  visibleInput: boolean = false;
  selectedIndex : number = 0;
  IMPORT_BATCH = IMPORT_BATCH
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
  output: any = {
    dlg: {},
    pt: [],
    db: [],
    fob: [],
    pt09: [],
    bbdo: [],
    pl1: [],
    pl2: [],
    pl3: [],
    pl4: [],
    vk11Pt: [],
    vk11Db: [],
    vk11Fob: [],
    vk11Tnpp: [],
    vk11Bb: [],
    summary: [],
  }
  headerId: any = '';
  constructor(
    private _service: CalculateDiscountService,
    private globalService: GlobalService,
    private message: NzMessageService,
    private router: Router,
    private route: ActivatedRoute,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Kết quả tính toán',
        path: 'calculate-discount/detail',
      },
    ])
    this.globalService.getLoading().subscribe((value) => {
      this.loading = value
    })
  }
  ngOnInit(): void {
    this.route.paramMap.subscribe({
      next: (params) => {
        const id = params.get('id')
        this.headerId = id
        this.getOutput(this.headerId);
      }
    })
    this._service.getInput(this.headerId).subscribe({
      next: (data) => {
        this.input = data;
     
      },
      error: (response) => {
        console.log(response)
      },
    })
    console.log(this.input)
  }
  getOutput(id: any) {
    this._service.getOutput(id).subscribe({
      next: (data) => {
        this.output = data;
        console.log(data)
      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  onClickTab(title: string, tab: number) {
    this.titleTab = title;
  }

  openInput() {
    this._service.getInput(this.headerId).subscribe({
      next: (data) => {
        this.input = data;
        this.visibleInput = true;
      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  onUpdateInput() {
    this._service.updateInput(this.input).subscribe({
      next: (data) => {
      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  close(): void {
    this.visibleInput = false;
  }
  reCalculate(){
    this.getOutput(this.headerId);
  }
  exportExcel(){
    this._service.exportExcel(this.headerId).subscribe({
      next: (data) => {
        console.log(data)
      },
      error: (response) => {
        console.log(response)
      },
    })
  }
}
