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
import { CuocVanChuyenFilter } from '../../models/master-data/cuoc-van-chuyen.model'
import { CuocVanChuyenListFilter } from '../../models/master-data/cuoc-van-chuyen-list.model'
import { CuocVanChuyenListService } from '../../services/master-data/cuoc-van-chuyen-list-service'
import { CuocVanChuyenService } from '../../services/master-data/cuoc-van-chuyen.service'

@Component({
  selector: 'app-cuoc-van-chuyen-list',
  standalone: true,
  imports: [ShareModule, NzButtonModule, NzIconModule, NzUploadModule],
  templateUrl: './cuoc-van-chuyen-list.component.html',
  styleUrl: './cuoc-van-chuyen-list.component.scss',
})
export class CuocVanChuyenListComponent {
  validateForm: FormGroup = this.fb.group({
    name: ['', [Validators.required]],
    isActive: [true, [Validators.required]],
  })

  isSubmit: boolean = false
  visible: boolean = false
  edit: boolean = false
  filter = new CuocVanChuyenListFilter()
  paginationResult = new PaginationResult()
  loading: boolean = false
  MASTER_DATA_MANAGEMENT = MASTER_DATA_MANAGEMENT

  constructor(
    private _service: CuocVanChuyenListService,
    private _serviceCuocVanChuyen: CuocVanChuyenService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private router: Router,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách cước vận chuyển',
        path: 'master-data/cuoc-van-chuyen',
      },
    ])
    this.globalService.getLoading().subscribe((value) => {
      this.loading = value
    })
  }
  fileList: NzUploadFile[] = []

  beforeUpload = (file: NzUploadFile): boolean => {
    file.originFileObj = file as any
    this.fileList = [file]
    return false
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
    this._service.searchCuocVanChuyen(this.filter).subscribe({
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

  navigate(data: any) {
    this.router.navigate([`/master-data/cuoc-van-chuyen/detail/${data}`])
  }

  submitForm(): void {
    this.isSubmit = true
    if (this.validateForm.valid && this.fileList[0]) {
      const formData = this.validateForm.getRawValue()
      formData.code = ''
      console.log(formData)
      this._service.createCuocVanChuyenList(formData).subscribe({
        next: (data) => {
          const formData = new FormData()
          formData.append('HeaderCode', data.code)
          if (this.fileList[0]?.originFileObj) {
            formData.append(
              'File',
              this.fileList[0].originFileObj as Blob,
              this.fileList[0].name,
            )
          }
          this._serviceCuocVanChuyen.createCuocVanChuyen(formData).subscribe({
            next: (data) => {
              this.fileList = []
              this.search()
            },
            error: (response) => {
              console.log(response)
            },
          })
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
    this.filter = new CuocVanChuyenFilter()
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
    this.validateForm.setValue({
      code: data.code,
      name: data.name,
      gap: data.gap,
      coefficient: data.coefficient,
      cuocVCBQ: data.cuocVCBQ,
      cpChungChuaCuocVC: data.cpChungChuaCuocVC,
      localCode: data.localCode,
      isActive: data.isActive,
      warehouseCode: data.warehouseCode,
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
