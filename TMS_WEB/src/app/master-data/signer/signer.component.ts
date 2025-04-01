import { Component } from '@angular/core'
import { ShareModule } from '../../shared/share-module'
import { GlobalService } from '../../services/global.service'
import { PaginationResult } from '../../models/base.model'
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms'
import { MASTER_DATA_MANAGEMENT } from '../../shared/constants'
import { NzButtonModule } from 'ng-zorro-antd/button'
import { Router } from '@angular/router'
import { NzIconModule } from 'ng-zorro-antd/icon'
import {
  NzUploadChangeParam,
  NzUploadFile,
  NzUploadModule,
} from 'ng-zorro-antd/upload'
import { SignerService } from '../../services/master-data/signer.service'
import { SignerFilter } from '../../models/master-data/signer.model'

@Component({
  selector: 'app-signer',
  standalone: true,
  imports: [ShareModule, NzButtonModule, NzIconModule, NzUploadModule],
  templateUrl: './signer.component.html',
  styleUrl: './signer.component.scss',
})
export class SignerComponent {
  validateForm: FormGroup = this.fb.group({
    code: [''],
    name: ['', [Validators.required]],
    position: ['', [Validators.required]],
    isActive: [true, [Validators.required]],
  })

  isSubmit: boolean = false
  visible: boolean = false
  edit: boolean = false
  filter = new SignerFilter()
  paginationResult = new PaginationResult()
  loading: boolean = false
  MASTER_DATA_MANAGEMENT = MASTER_DATA_MANAGEMENT

  constructor(
    private _service: SignerService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private router: Router,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách người ký',
        path: 'master-data/signer',
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
    this._service.searchSigners(this.filter).subscribe({
      next: (data) => {
        this.paginationResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  isCodeExist(code: string): boolean {
    return this.paginationResult.data?.some(
      (market: any) => market.code === code,
    )
  }

  submitForm(): void {
    this.isSubmit = true
    if (this.validateForm.valid) {
      const formData = this.validateForm.getRawValue()
      formData.code = ''
      console.log(formData)
      this._service.createSigner(formData).subscribe({
        next: (data) => {
          this.search()
        },
        error: (response) => {
          console.log(response)
        },
      })
    } else {
      Object.values(this.validateForm.controls).forEach((control) => {
        if (control.invalid) {
          control.markAsDirty()
          control.updateValueAndValidity({ onlySelf: true })
        }
      })
    }
  }

  updateSigner(): void {
    this.isSubmit = true
    if (this.validateForm.valid) {
      const formData = this.validateForm.getRawValue()
      this._service.updateSigner(formData).subscribe({
        next: (data) => {
          this.search()
        },
        error: (response) => {
          console.log(response)
        },
      })
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
    this.filter = new SignerFilter()
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

  openEdit(data: any) {
    console.log(data)
    this.validateForm.setValue({
      code: data.code,
      name: data.name,
      position: data.position,
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
