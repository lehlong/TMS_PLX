import { Component } from '@angular/core'
import { ShareModule } from '../../shared/share-module'
import { GlobalService } from '../../services/global.service'
import { PaginationResult } from '../../models/base.model'
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms'
import { MARKET_COMPETITOR_RIGHTS, MASTER_DATA_MANAGEMENT } from '../../shared/constants'
import { NzMessageService } from 'ng-zorro-antd/message'
import { MarketCompetitorFilter } from '../../models/master-data/market-competitor.model'
import { MarketCompetitorService } from '../../services/master-data/market-competitor.service'
import { MarketService } from '../../services/master-data/market.service'
import { CompetitorService } from '../../services/master-data/competitor.service'

@Component({
  selector: 'app-market',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './market-competitor.component.html',
  styleUrl: './market-competitor.component.scss',
})

export class MarketCompetitorComponent {
  validateForm: FormGroup = this.fb.group({
    code: [''],
    gap: ['', [Validators.required]],
    marketCode: ['', [Validators.required]],
    competitorCode: ['', [Validators.required]],
    isActive: [true, [Validators.required]],
  })

  isSubmit: boolean = false
  visible: boolean = false
  edit: boolean = false
  filter = new MarketCompetitorFilter()
  paginationResult = new PaginationResult()
  loading: boolean = false
  marketResult: any[] = []
  competitorResult: any[] = []
  MASTER_DATA_MANAGEMENT = MASTER_DATA_MANAGEMENT

  constructor(
    private _service: MarketCompetitorService,
    private _marketService: MarketService,
    private _competitorService: CompetitorService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private message: NzMessageService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách thị trường đối thủ',
        path: 'master-data/market-competitor',
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
    this.getAllMarket()
    this.getAllCompetitor()
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
    this._service.searchmarket(this.filter).subscribe({
      next: (data) => {
        this.paginationResult = data
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

  getAllCompetitor() {
    this.isSubmit = false
    this._competitorService.getall().subscribe({
      next: (data) => {
        this.competitorResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  exportExcel() {
    return this._service
      .exportExcelmarket(this.filter)
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

  isCodeExist(code: string, competitorCode:string): boolean {
    return this.paginationResult.data?.some((market: any) => market.marketCode === code && market.competitorCode ===  competitorCode)
  }
  // && market.competitorCode ===
  submitForm(): void {
    this.isSubmit = true
    if (this.validateForm.valid) {
      const formData = this.validateForm.getRawValue()
      if (this.edit) {

        if(this.paginationResult.data?.find((market: any) => market.code === formData.code && market.competitorCode !=  formData.competitorCode)){
          this.message.error(
            `Mã thị trường ${formData.marketCode} đã được tạo cho đối thủ ${formData.competitorCode} trước đó`,
          )
          return
        }        
        this._service.updatemarket(formData).subscribe({
          next: (data) => {
            this.search()
          },
          error: (response) => {
            console.log(response)
          },
        })
      } else {
        if (this.isCodeExist(formData.marketCode,formData.competitorCode)) {
          this.message.error(
            `Mã thị trường ${formData.marketCode} đã được tạo cho đối thủ ${formData.competitorCode} trước đó`,
          )
          return
        }
        this._service.createmarket(formData).subscribe({
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
    this.filter = new MarketCompetitorFilter()
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
    this._service.deletemarket(code).subscribe({
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
      gap: data.gap,
      marketCode: data.marketCode,
      isActive: data.isActive,
      competitorCode: data.competitorCode
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
