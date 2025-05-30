import { Component } from '@angular/core'
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms'
import { NzMessageService } from 'ng-zorro-antd/message'
import { PaginationResult } from '../../../models/base.model'
import { GlobalService } from '../../../services/global.service'
import { MASTER_DATA_MANAGEMENT } from '../../../shared/constants'
import { ShareModule } from '../../../shared/share-module'
import { LocalService } from '../../../services/master-data/local.service'
import { MarketService } from '../../../services/master-data/market.service'
import { CustomerPtFilter } from '../../../models/master-data/customer-pt.model'
import { TermOfPaymentService } from '../../../services/master-data/term-of-payment.service'
import { CustomerFobService } from '../../../services/master-data/customer-fob.service'
@Component({
  selector: 'fob',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './customer-fob.component.html',
  styleUrl: './customer-fob.component.scss',
})
export class CustomerFobComponent {
  validateForm: FormGroup = this.fb.group({
    code: ['', [Validators.required]],
    name: ['', [Validators.required]],
    marketCode: [''],
    localCode: [''],
    cuLyBq: [0],
    cvcbq: [0],
    cpccvc: [0],
    ckXang: [0],
    ckDau: [0],
    htcvc: [0],
    httVb1370: [0],
    ckv2: [0],
    adrress: [''],
    thtt: [''],
    lvnh: [0],
    order: [0],
    phuongThuc: [''],
    isActive: [true, [Validators.required]],
    lamTronDacBiet: [false, [Validators.required]],
  })

  isSubmit: boolean = false
  visible: boolean = false
  edit: boolean = false
  filter = new CustomerPtFilter()
  paginationResult = new PaginationResult()
  loading: boolean = false
  MASTER_DATA_MANAGEMENT = MASTER_DATA_MANAGEMENT
  lstType: any[] = []
  data: any = []
  localResult: any = []
  marketResult: any = []

  thttLst: any = []


    constructor(
    private _service: CustomerFobService,
    private _thttService: TermOfPaymentService,
    private _localService: LocalService,
    private _marketService: MarketService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private message: NzMessageService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách khách hàng',
        path: 'master-data/customer/fob',
      },
    ])
    this.globalService.getLoading().subscribe((value) => {
      this.loading = value
    })
  }

  ngOnDestroy() {
    this.globalService.setBreadcrumb([])
  }

  ngOnInit(): void {
    this.search()
    this.getAllLocal()
    this.getAllThtt()
    this.getAllMarket()
    this.lstType = [
      { code: 'X', name: 'Xăng' },
      { code: 'D', name: 'Dầu' }
    ]
  }

  onSortChange(name: string, value: any) {
    this.filter = {
      ...this.filter,
      SortColumn: name,
      IsDescending: value === 'descend',
    }
    this.search()
  }

  search() {
    this.isSubmit = false
    this._service.searchCustomerFob(this.filter).subscribe({
      next: (data) => {
        this.paginationResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }


  getAllLocal() {
    this.isSubmit = false
    this._localService.getall().subscribe({
      next: (data) => {
        this.localResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  getAllThtt() {
    this.isSubmit = false
    this._thttService.getall().subscribe({
      next: (data) => {
        this.thttLst = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  getAllMarket() {
    this.isSubmit = false
    this._marketService.getall().subscribe({
      next: (data) => {
        this.marketResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  exportExcel() {
    return this._service
      .exportExcelCustomerFob(this.filter)
      .subscribe((result: Blob) => {
        const blob = new Blob([result], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        const url = window.URL.createObjectURL(blob)
        var anchor = document.createElement('a')
        anchor.download = 'danh-sach-dia-phuong.xlsx'
        anchor.href = url
        anchor.click()
      })
  }
  isCodeExist(code: string): boolean {
    return this.paginationResult.data?.some((local: any) => local.code === code)
  }
  submitForm(): void {
    this.isSubmit = true
    // if (this.validateForm.valid) {
    const formData = this.validateForm.getRawValue()
    if (this.edit) {
      this._service.updateCustomerFob(formData).subscribe({
        next: (data) => {
          this.search()
        },
        error: (response) => {
          console.log(response)
        },
      })
    } else {
      if (this.isCodeExist(formData.code)) {
        this.message.error(
          `Mã khu vục ${formData.code} đã tồn tại, vui lòng nhập lại`,
        )
        return
      }
      this._service.createCustomerFob(formData).subscribe({
        next: (data) => {
          this.search()
        },
        error: (response) => {
          console.log(response)
        },
      })
    }
  }

  close() {
    this.visible = false
    this.resetForm()
  }

  reset() {
    this.filter = new CustomerPtFilter()
    this.search()
  }

  openCreate() {
    this.edit = false
    this.visible = true
  }

  resetForm() {
    this.validateForm.reset()
    this.isSubmit = false
  }

  deleteItem(code: string | number) {
    this._service.deleteCustomerFob(code).subscribe({
      next: (data) => {
        this.search()
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  openEdit(data: any) {
    this.validateForm.setValue({
      code: data.code,
      name: data.name,
      localCode: data.localCode,
      marketCode: data.marketCode,
      ckDau: data.ckDau,
      ckXang: data.ckXang,
      ckv2: data.ckv2,
      cpccvc: data.cpccvc,
      cuLyBq: data.cuLyBq,
      cvcbq: data.cvcbq,
      htcvc: data.htcvc,
      httVb1370: data.httVb1370,
      isActive: data.isActive,
      lvnh: data.lvnh,
      order: data.order,
      phuongThuc: data.phuongThuc,
      thtt: data.thtt,
      adrress: data.adrress,
      lamTronDacBiet: data.lamTronDacBiet,

    })
    setTimeout(() => {
      this.edit = true
      this.visible = true
    }, 200)
  }

  pageSizeChange(size: number): void {
    this.filter.currentPage = 1
    this.filter.pageSize = size
    this.search()
  }

  pageIndexChange(index: number): void {
    this.filter.currentPage = index
    this.search()
  }
}
