import { Component } from '@angular/core'
import { ShareModule } from '../../shared/share-module'
import { GlobalService } from '../../services/global.service'
import { PaginationResult } from '../../models/base.model'
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms'
import { DELIVERY_POINT_RIGHTS, MASTER_DATA_MANAGEMENT } from '../../shared/constants'
import { NzMessageService } from 'ng-zorro-antd/message'
import { DeliveryPointFilter } from '../../models/master-data/delivery-point.model'
import { DeliveryPointService } from '../../services/master-data/delivery-point.service'
import { CustomerService } from '../../services/master-data/customer.service'
@Component({
  selector: 'app-local',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './delivery-point.component.html',
  styleUrl: './delivery-point.component.scss',
})
export class DeliveryPointComponent {
  validateForm: FormGroup = this.fb.group({
    code: [''],
    name: ['', [Validators.required]],
    cuocVcBq: [''],
    isActive: [true, [Validators.required]],
  })

  isSubmit: boolean = false
  visible: boolean = false
  edit: boolean = false
  filter = new DeliveryPointFilter()
  paginationResult = new PaginationResult()
  loading: boolean = false
  customerResult: any[] = []
  MASTER_DATA_MANAGEMENT = MASTER_DATA_MANAGEMENT


  constructor(
    private _service: DeliveryPointService,
    private _customerService: CustomerService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private message: NzMessageService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách điểm giao hàng',
        path: 'master-data/delivery-point',
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
    this.getAllCustomer()

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
    this._service.searchDeliveryPoint(this.filter).subscribe({
      next: (data) => {
        this.paginationResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  getAllCustomer() {
    this.isSubmit = false
    this._customerService.getall().subscribe({
      next: (data) => {
        this.customerResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  exportExcel() {
    return this._service
      .exportExcelDeliveryPoint(this.filter)
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
    if (this.validateForm.valid) {
      const formData = this.validateForm.getRawValue()
      if (this.edit) {
        this._service.updateDeliveryPoint(formData).subscribe({
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
            `Mã ${formData.code} đã tồn tại, vui lòng nhập lại`,
          )
          return
        }
        this._service.createDeliveryPoint(formData).subscribe({
          next: (data) => {
            this.search()
          },
          error: (response) => {
            console.log(response)
          },
        })
      }
    } else {
      Object.values(this.validateForm.controls).forEach((control) => {
        if (control.invalid) {
          control.markAsDirty()
          control.updateValueAndValidity({ onlySelf: true })
        }
      })
    }
  }

  close() {
    this.visible = false
    this.resetForm()
  }

  reset() {
    this.filter = new DeliveryPointFilter()
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
    this._service.deleteDeliveryPoint(code).subscribe({
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
      cuocVcBq: data.cuocVcBq,
      isActive: data.isActive,
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
