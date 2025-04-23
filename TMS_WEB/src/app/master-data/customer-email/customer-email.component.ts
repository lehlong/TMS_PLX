import { Component } from '@angular/core'
import { ShareModule } from '../../shared/share-module'
import { GlobalService } from '../../services/global.service'
import { PaginationResult } from '../../models/base.model'
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms'
import { MASTER_DATA_MANAGEMENT } from '../../shared/constants'
import { NzButtonModule } from 'ng-zorro-antd/button'
import { NzIconModule } from 'ng-zorro-antd/icon'
import { SignerFilter } from '../../models/master-data/signer.model'
import { CalculateDiscountService } from '../../services/calculate-discount/calculate-discount.service'
import { CustomerEmailService } from '../../services/master-data/customer-email.service'

@Component({
  selector: 'app-customer-email',
  standalone: true,
  imports: [ShareModule, NzButtonModule, NzIconModule],
  templateUrl: './customer-email.component.html',
  styleUrl: './customer-email.component.scss'
})
export class CustomerEmailComponent {
isSubmit: boolean = false
  visible: boolean = false
  edit: boolean = false
  filter = new SignerFilter()
  paginationResult = new PaginationResult()
  loading: boolean = false
  MASTER_DATA_MANAGEMENT = MASTER_DATA_MANAGEMENT
  lstEmail:any[] = []
  constructor(
    private _service: CustomerEmailService,
    private _CalculateDiscountservice: CalculateDiscountService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách Email',
        path: 'master-data/customer-phone',
      },
    ])
    this.globalService.getLoading().subscribe((value) => {
      this.loading = value
    })
  }
  validateForm: FormGroup = this.fb.group({
    code:'',
    customerCode: ['', [Validators.required]],
    email: ['', [Validators.required]],
    isActive: [true, [Validators.required]],
  })
  get customerCode() {
    return this.validateForm.get('customerCode');
  }
  lstCustomer:any[] =[]
  ngOnDestroy() {
    this.globalService.setBreadcrumb([])
  }

  ngOnInit(): void {
    this.search()
    this.getAllInputCustomer()
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
    this._service.search(this.filter).subscribe({
      next: (data) => {
        this.paginationResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  submitForm(): void {
    this.isSubmit = true
    if (this.validateForm.valid) {
      const formData = this.validateForm.getRawValue()
      formData.code = ''
      this._service.create(formData).subscribe({
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

  getAllInputCustomer():void{
    this._CalculateDiscountservice.GetAllInputCustomer().subscribe({
      next: (data) => {
        this.lstCustomer = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  updateCustomerEmail(): void {
    this.isSubmit = true
    if (this.validateForm.valid) {
      const formData = this.validateForm.getRawValue()
      this._service.update(formData).subscribe({
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
    this.validateForm.setValue({
      code: data.code,
      customerCode: data.customerCode,
      email: data.email,
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
  getNameByCode(code: string): string {
    const customer = this.lstCustomer.find(x => x.code === code);
    return customer ? customer.name : ''; // Trả về tên nếu có, nếu không có thì trả về chuỗi rỗng.
  }
}
