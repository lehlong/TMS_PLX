import { Component } from '@angular/core'
import { ShareModule } from '../../shared/share-module'
import { FormGroup, NonNullableFormBuilder, Validators } from '@angular/forms'
import { AccountTypeFilter } from '../../models/master-data/account-type.model'
import { PaginationResult } from '../../models/base.model'
import { ACCOUNTTYPE_RIGHTS } from '../../shared/constants'
import { AccountTypeService } from '../../services/master-data/account-type.service'
import { GlobalService } from '../../services/global.service'
import { NzMessageService } from 'ng-zorro-antd/message'

@Component({
  selector: 'app-account-type',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './account-type.component.html',
  styleUrl: './account-type.component.scss',
})
export class AccountTypeComponent {
  validateForm: FormGroup = this.fb.group({
    id: ['', [Validators.required]],
    name: ['', [Validators.required]],
    isActive: [true, [Validators.required]],
  })
  isSubmit: boolean = false
  visible: boolean = false
  edit: boolean = false
  filter = new AccountTypeFilter()
  paginationResult = new PaginationResult()
  loading: boolean = false
 
  constructor(
    private _service: AccountTypeService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private message: NzMessageService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách kiểu người dùng',
        path: 'master-data/account-type',
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
    this._service.searchAccountType(this.filter).subscribe({
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
      .exportExcelAccountType(this.filter)
      .subscribe((result: Blob) => {
        const blob = new Blob([result], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        const url = window.URL.createObjectURL(blob)
        var anchor = document.createElement('a')
        anchor.download = 'danh-sach-kieu-nguoi-dung.xlsx'
        anchor.href = url
        anchor.click()
      })
  }
  isCodeExist(code: string): boolean {
    return this.paginationResult.data?.some(
      (accType: any) => accType.id === code,
    )
  }
  submitForm(): void {
    this.isSubmit = true
    if (this.validateForm.valid) {
      if (this.edit) {
        this._service
          .updateAccountType(this.validateForm.getRawValue())
          .subscribe({
            next: (data) => {
              this.search()
            },
            error: (response) => {
              console.log(response)
            },
          })
      } else {
        const formData = this.validateForm.getRawValue()
        if (this.isCodeExist(formData.id)) {
          this.message.error(
            `Mã kiểu người dùng ${formData.id} đã tồn tại, vui lòng nhập lại`,
          )
          return
        }
        this._service
          .createAccountType(this.validateForm.getRawValue())
          .subscribe({
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
    this.filter = new AccountTypeFilter()
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

  deleteItem(id: string) {
    this._service.deleteAccountType(id).subscribe({
      next: (data) => {
        this.search()
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  openEdit(data: { id: string; name: number; isActive: boolean }) {
    this.validateForm.setValue({
      id: data.id,
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
