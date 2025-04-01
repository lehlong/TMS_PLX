import { Component, OnInit, output } from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { NzMessageService } from 'ng-zorro-antd/message';
import { CalculateDiscountService } from '../../services/calculate-discount/calculate-discount.service';
import { GlobalService } from '../../services/global.service';
import { ShareModule } from '../../shared/share-module';
import { environment } from '../../../environments/environment.prod'
import {
  CALCULATE_RESULT_RIGHT,
  IMPORT_BATCH,
} from '../../shared/constants/access-right.constants'
import { SignerService } from '../../services/master-data/signer.service';
import Swal from 'sweetalert2';

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
  isVisibleStatus: boolean = false;
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
  statusModel = {
    title: '',
    des: '',
    value: '',
  }
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
  checked = false
  headerId: any = '';
  signerResult: any[] = []
  isVisibleLstTrinhKy: boolean = false
  lstTrinhKyChecked: any[] = []
  lstTrinhKy: any[] = [
    {
      code: 'CongDienKKGiaBanLe',
      name: 'Công Điện Kiểm Kê Giá Bán Lẻ',
      status: true,
    },
    { code: 'MucGiamGiaNQTM', name: 'Mức Giảm Giá NQTM', status: false },
    { code: 'QDGBanBuon', name: 'Quyết Định Giá bán Buôn', status: false },
    { code: 'QDGBanLe', name: 'Quyết Định Giá Bán lẻ', status: true },
    { code: 'QDGCtyPTS', name: 'Quyết Định Công Ty PTS', status: false },
    { code: 'QDGNoiDung', name: 'Quyết Giá Nội Dụng', status: false },
    { code: 'ToTrinh', name: 'Tờ Trình', status: false },
    { code: 'KeKhaiGia', name: 'Kê Khai Giá', status: true },
    { code: 'KeKhaiGiaChiTiet', name: 'Kê Khai Giá Chi Tiết', status: true },
  ]
  constructor(
    private _service: CalculateDiscountService,
    private globalService: GlobalService,
    private message: NzMessageService,
    private router: Router,
    private route: ActivatedRoute,
    private _signerService: SignerService,
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

  exportWordTrinhKy() {
    
    this.isVisibleLstTrinhKy = !this.isVisibleLstTrinhKy
    console.log(this.isVisibleLstTrinhKy);
  }
  updateTrinhKyCheckedSet(code: any, checked: boolean): void {
    if (checked) {
      this.lstTrinhKyChecked.push(code)
    } else {
      this.lstTrinhKyChecked = this.lstTrinhKyChecked.filter((x) => x != code)
    }
  }

  onItemTrinhKyChecked(code: String, checked: boolean): void {
    this.updateTrinhKyCheckedSet(code, checked)
    console.log(this.lstTrinhKyChecked)
  }

  onAllCheckedLstTrinhKy(value: boolean): void {
    this.lstTrinhKyChecked = []
    if (value) {
      this.lstTrinhKy.forEach((i) => {
        this.lstTrinhKyChecked.push(i.code)
      })
    } else {
      this.lstTrinhKyChecked = []
    }
  }
  confirmExportWordTrinhKy() {
    if (this.lstTrinhKyChecked.length == 0) {
      this.message.create('warning', 'Vui lòng chọn trình ky xuất ra file')
      return
    } else {
      this._service
        .ExportWordTrinhKy(this.lstTrinhKyChecked, this.headerId)
        .subscribe({
          next: (data) => {
            this.lstTrinhKyChecked = []
            var a = document.createElement('a')
            a.href = environment.apiUrl + data
            a.target = '_blank'
            a.click()
            a.remove()
          },
          error: (err) => {
            console.log(err)
          },
        })
      this.lstTrinhKyChecked = []
    }
  }

  onClickTab(title: string, tab: number) {
    this.titleTab = title;
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
    // this.input.status.code = status
    // this.isVisibleStatus = true
    Swal.fire({
      title: this.statusModel.title,
      text: this.statusModel.des,
      input: 'text',
      inputPlaceholder: 'Ý kiến',
      icon: 'question',
      showCancelButton: true,
      confirmButtonText: 'Đồng ý',
      cancelButtonText: 'Hủy'
    }).then((result) => {
      if (result.isConfirmed) {
        console.log(result.value)
      }
    });

  }

  getAllSigner() {
    this._signerService.getall().subscribe({
      next: (data) => {
        this.signerResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
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
    this.getAllSigner()
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
