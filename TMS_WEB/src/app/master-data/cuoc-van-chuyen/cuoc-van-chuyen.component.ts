import { Component } from '@angular/core'
import { ShareModule } from '../../shared/share-module'
import { GlobalService } from '../../services/global.service'
import { PaginationResult } from '../../models/base.model'
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms'
import { CUOC_VAN_CHUYEN_RIGHTS } from '../../shared/constants'
import { NzMessageService } from 'ng-zorro-antd/message'
import { NzButtonModule } from 'ng-zorro-antd/button';
import { NzIconModule } from 'ng-zorro-antd/icon';
import { NzUploadChangeParam, NzUploadFile, NzUploadModule } from 'ng-zorro-antd/upload';
import { CuocVanChuyenService } from '../../services/master-data/cuoc-van-chuyen.service'
import { CuocVanChuyenFilter } from '../../models/master-data/cuoc-van-chuyen.model'
import { ActivatedRoute } from '@angular/router'

  @Component({
    selector: 'app-cuoc-van-chuyen',
    standalone: true,
    imports: [ShareModule,NzButtonModule, NzIconModule, NzUploadModule],
    templateUrl: './cuoc-van-chuyen.component.html',
    styleUrl: './cuoc-van-chuyen.component.scss'
  })
  export class CuocVanChuyenComponent {
  validateForm: FormGroup = this.fb.group({
      code: ['', [Validators.required]],
      name: ['', [Validators.required]],
      gap: ['', [Validators.required]],
      coefficient: [1.1, [Validators.required]],
      cuocVCBQ: ['', [Validators.required]],
      cpChungChuaCuocVC: ['', [Validators.required]],
      localCode: ['', [Validators.required]],
      warehouseCode: [''],
      isActive: [true, [Validators.required]],
    })

    isSubmit: boolean = false
    visible: boolean = false
    edit: boolean = false
    filter = new CuocVanChuyenFilter()
    paginationResult = new PaginationResult()
    loading: boolean = false
    CUOC_VAN_CHUYEN_RIGHTS = CUOC_VAN_CHUYEN_RIGHTS

    constructor(
      private _service: CuocVanChuyenService,
      private fb: NonNullableFormBuilder,
      private globalService: GlobalService,
      private message: NzMessageService,
      private route: ActivatedRoute
    ) {
      this.globalService.setBreadcrumb([
        {
          name: 'Chi tiết cước vận chuyển',
          path: 'master-data/cuoc-van-chuyen',
        },
      ])
      this.globalService.getLoading().subscribe((value) => {
        this.loading = value
      })
    }
    fileList: NzUploadFile[] = [
    ];

    handleChange(info: NzUploadChangeParam): void {
      let fileList = [...info.fileList];

      fileList = fileList.slice(-1);

      console.log(this.fileList);
      
      this.fileList = fileList;
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
      this.route.paramMap.subscribe((params) => {
        const code = params.get('code');
        if (code) {
          this._service.searchCuocVanChuyen(this.filter,code).subscribe({
            next: (data) => {
              this.paginationResult = data
            },
            error: (response) => {
              console.log(response)
            },
          })
        }
      });
      
    }

    isCodeExist(code: string): boolean {
      return this.paginationResult.data?.some((market: any) => market.code === code)
    }

    submitForm(): void {
      this.isSubmit = true
      // if (this.validateForm.valid) {
      //   const formData = this.validateForm.getRawValue()
      //   console.log(formData);

      //   if (this.edit) {
      //   }
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

    // deleteItem(code: string | number) {
    //   this._service.deletemarket(code).subscribe({
    //     next: (data) => {
    //       this.search()
    //     },
    //     error: (response) => {
    //       console.log(response)
    //     },
    //   })
    // }

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
        warehouseCode: data.warehouseCode
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
