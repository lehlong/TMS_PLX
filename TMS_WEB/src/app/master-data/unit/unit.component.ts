import { Component } from '@angular/core'
import { ShareModule } from '../../shared/share-module'
import { UnitFilter } from '../../models/master-data/unit.model'
import { GlobalService } from '../../services/global.service'
import { UnitService } from '../../services/master-data/unit.service'
import { PaginationResult } from '../../models/base.model'
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms'
import { UNIT_RIGHTS } from '../../shared/constants'
import { NzMessageService } from 'ng-zorro-antd/message'
@Component({
  selector: 'app-unit',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './unit.component.html',
  styleUrl: './unit.component.scss',
})
export class UnitComponent {
  validateForm: FormGroup = this.fb.group({
    code: ['', [Validators.required]],
    name: ['', [Validators.required]],
    isActive: [true, [Validators.required]],
  })

  isSubmit: boolean = false
  visible: boolean = false
  edit: boolean = false
  filter = new UnitFilter()
  paginationResult = new PaginationResult()
  loading: boolean = false
  UNIT_RIGHTS = UNIT_RIGHTS
  constructor(
    private _service: UnitService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private message: NzMessageService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách đơn vị tính',
        path: 'master-data/unit',
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
    this._service.searchUnit(this.filter).subscribe({
      next: (data) => {
        this.paginationResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  exportExcel() {
    return this._service
      .exportExcelUnit(this.filter)
      .subscribe((result: Blob) => {
        const blob = new Blob([result], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        const url = window.URL.createObjectURL(blob)
        var anchor = document.createElement('a')
        anchor.download = 'danh-sach-don-vi-tinh.xlsx'
        anchor.href = url
        anchor.click()
      })
  }
  isCodeExist(code: string): boolean {
    return this.paginationResult.data?.some((unit: any) => unit.code === code)
  }
  submitForm(): void {
    this.isSubmit = true
    if (this.validateForm.valid) {
      const formData = this.validateForm.getRawValue()
      if (this.edit) {
        this._service.updateUnit(formData).subscribe({
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
            `Mã đơn vị ${formData.code} đã tồn tại, vui lòng nhập lại`,
          )
          return
        }
        this._service.createUnit(formData).subscribe({
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
    this.filter = new UnitFilter()
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
    this._service.deleteUnit(code).subscribe({
      next: (data) => {
        this.search()
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  openEdit(data: { code: string; name: string; isActive: boolean }) {
    this.validateForm.setValue({
      code: data.code,
      name: data.name,
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
