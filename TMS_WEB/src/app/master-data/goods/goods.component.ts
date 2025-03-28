import { Component } from '@angular/core'
import { ShareModule } from '../../shared/share-module'
import { LocalFilter } from '../../models/master-data/local.model'
import { GlobalService } from '../../services/global.service'
import { LocalService } from '../../services/master-data/local.service'
import { PaginationResult } from '../../models/base.model'
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms'
import { LOCAL_RIGHTS, GOODS_RIGHTS, MASTER_DATA_MANAGEMENT } from '../../shared/constants'
import { NzMessageService } from 'ng-zorro-antd/message'
import { GoodsFilter } from '../../models/master-data/goods.model'
import { GoodsService } from '../../services/master-data/goods.service'
@Component({
  selector: 'app-local',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './goods.component.html',
  styleUrl: './goods.component.scss',
})
export class GoodsComponent {
  validateForm: FormGroup = this.fb.group({
    code: ['', [Validators.required]],
    name: ['', [Validators.required]],
    type: [''],
    thueBvmt: ['', [Validators.required]],
    vfc: ['', [Validators.required]],
    mtsV1: ['', [Validators.required]],
    order: ['', [Validators.required]],
    createDate: [new Date(), [Validators.required]],
    isActive: [true, [Validators.required]],
  })

  isSubmit: boolean = false
  visible: boolean = false
  edit: boolean = false
  filter = new GoodsFilter()
  paginationResult = new PaginationResult()
  loading: boolean = false
  MASTER_DATA_MANAGEMENT = MASTER_DATA_MANAGEMENT
  lstType: any[] = []

  constructor(
    private _service: GoodsService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private message: NzMessageService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách mặt hàng',
        path: 'master-data/goods',
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
    this.lstType = [
      {code:'X', name: 'Xăng'},
      {code:'D', name: 'Dầu'}
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
    this._service.searchGoods(this.filter).subscribe({
      next: (data) => {
        this.paginationResult = data
        console.log(this.paginationResult);

      },
      error: (response) => {
        console.log(response)
      },
    })
  }



  exportExcel() {
    return this._service
      .exportExcelGoods(this.filter)
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
        console.log(formData);

        this._service.updateGoods(formData).subscribe({
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
        this._service.createGoods(formData).subscribe({
          next: (data) => {
            this.search()
          },
          error: (response) => {
            console.log(response)
          },
        })
      }
    // } else {
    //   Object.values(this.validateForm.controls).forEach((control) => {
    //     if (control.invalid) {
    //       control.markAsDirty()
    //       control.updateValueAndValidity({ onlySelf: true })
    //     }
    //   })
    // }
  }

  close() {
    this.visible = false
    this.resetForm()
  }

  reset() {
    this.filter = new GoodsFilter()
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
    this._service.deleteGoods(code).subscribe({
      next: (data) => {
        this.search()
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  openEdit(data: any) {
    console.log(data)
    this.validateForm.setValue({
      code: data.code,
      name: data.name,
      type: data.type,
      thueBvmt: data.thueBvmt,
      vfc: data.vfc,
      mtsV1: data.mtsV1,
      order: data.order,
      createDate: data.createDate,
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
