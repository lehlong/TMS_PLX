import { Component } from '@angular/core'
import { ShareModule } from '../../shared/share-module'
import { GlobalService } from '../../services/global.service'
import { LocalService } from '../../services/master-data/local.service'
import { PaginationResult } from '../../models/base.model'
import {
  FormGroup,
  Validators,
  NonNullableFormBuilder,
  FormsModule,
} from '@angular/forms'
import { MASTER_DATA_MANAGEMENT } from '../../shared/constants'
import { NzMessageService } from 'ng-zorro-antd/message'
import { CustomerFilter } from '../../models/master-data/customer.model'
import { CustomerService } from '../../services/master-data/customer.service'
import { MarketService } from '../../services/master-data/market.service'
import { SalesMethodService } from '../../services/master-data/sales-method.service'
import { CustomerTypeService } from '../../services/master-data/customer-type.service'
import { NzDividerModule } from 'ng-zorro-antd/divider'
import { NzIconModule } from 'ng-zorro-antd/icon'
import { NzInputModule } from 'ng-zorro-antd/input'
import { NzSelectModule } from 'ng-zorro-antd/select'
import { CustomerContactService } from '../../services/master-data/customer-contact.service'
import { NzDropDownModule } from 'ng-zorro-antd/dropdown'
@Component({
  selector: 'app-local',
  standalone: true,
  imports: [
    ShareModule,
    FormsModule,
    NzDividerModule,
    NzIconModule,
    NzInputModule,
    NzSelectModule,
    NzDropDownModule,
  ],
  templateUrl: './customer.component.html',
  styleUrl: './customer.component.scss',
})
export class CustomerComponent {
  validateForm: FormGroup = this.fb.group({
    code: [''],
    name: ['', [Validators.required]],
    address: [''],
    paymentTerm: [''],
    gap: [0],
    cuocVcBq: [0],
    mgglhXang: [0],
    mgglhDau: [0],
    buyInfo: [''],
    bankLoanInterest: [0],
    salesMethodCode: [''],
    customerTypeCode: [''],
    localCode: [''],
    marketCode: [''],
    fob:[0],
    isActive: [true, [Validators.required]],
  })
  listOfPhone: {
    type: string
    code: string
    customer_Code: string
    value: string
    isActive: boolean
  }[] = []
  listOfEmail: {
    type: string
    code: string
    customer_Code: string
    value: string
    isActive: boolean
  }[] = []

  index = 0

  visiblePhone: any[] = [];
  visibleEmail: any[] = [];

  visibleChange(value: boolean,i:number,isPhone:boolean): void {
   isPhone?this.visiblePhone[i] = value:this.visibleEmail[i] = value;
  }

  contactList: any[] = []
  listOfPhoneOptions: { [key: string]: string[] } = {} // Lưu danh sách số điện thoại theo customerCode
  listOfEmailOptions: { [key: string]: string[] } = {} // Lưu danh sách email theo customerCode
  readonly nzFilterOption = (): boolean => true
  selectedPhones: { [key: string]: string } = {}
  selectedEmails: { [key: string]: string } = {}

  searchContactList(
    contactValue: string,
    type: 'phone' | 'email' | 'both',
    customerCode: string,
  ): void {
    if (type === 'phone' || type === 'both') {
      this.listOfPhoneOptions[customerCode] = this.contactList
        .filter(
          (item) =>
            item.type === 'phone' && item.customer_Code === customerCode,
        )
        .map((item) => item.value)
        .filter(
          (value) =>
            !contactValue ||
            value.toLowerCase().includes(contactValue.toLowerCase()),
        )
    } else if (type === 'email' || type === 'both') {
      this.listOfEmailOptions[customerCode] = this.contactList
        .filter(
          (item) =>
            item.type === 'email' && item.customer_Code === customerCode,
        )
        .map((item) => item.value)
        .filter(
          (value) =>
            !contactValue ||
            value.toLowerCase().includes(contactValue.toLowerCase()),
        )
    }
  }

  updatePhoneValue(index: number, newValue: string): void {
    this.listOfPhone[index] = { ...this.listOfPhone[index], value: newValue }
  }

  addItem(input: HTMLInputElement, isPhone: boolean): void {
    const value = input.value.trim()

    if (
      isPhone &&
      value &&
      !this.listOfPhone.some((item) => item.value === value)
    ) {
      this.listOfPhone.push({
        type: 'phone',
        code: ``,
        customer_Code: '',
        value: value,
        isActive: true,
      })

      // Cập nhật lại mảng để Angular nhận diện thay đổi
      this.listOfPhone = [...this.listOfPhone]

      // Xóa nội dung input sau khi thêm
      input.value = ''
    } else if (
      !isPhone &&
      value &&
      !this.listOfEmail.some((item) => item.value === value)
    ) {
      this.listOfEmail.push({
        type: 'email',
        code: ``,
        customer_Code: '',
        value: value,
        isActive: true,
      })

      // Cập nhật lại mảng để Angular nhận diện thay đổi
      this.listOfEmail = [...this.listOfEmail]

      // Xóa nội dung input sau khi thêm
      input.value = ''
    }
  }

  getFirstActiveContact(isPhone:boolean): string {
    if(isPhone){
      const activePhone = this.listOfPhone.find(item => item.isActive === true);
      return activePhone ? activePhone.value : "Danh sách số điện thoại";
    }
   else{
    const activeEmail = this.listOfEmail.find(item => item.isActive === true);
    return activeEmail ? activeEmail.value : "Danh sách email";
   }
  }

  removeItem(index: number, isPhone: boolean, isUpdate: boolean): void {
    if (isUpdate) {
      if (isPhone) {
        this.listOfPhone[index].isActive = false
      } else {
        this.listOfEmail[index].isActive = false
      }
    } else {
      if (isPhone) {
        this.listOfPhone.splice(index, 1)
      } else {
        this.listOfEmail.splice(index, 1)
      }
    }
    isPhone?this.visiblePhone[index] = false:this.visibleEmail[index]=false;
  }

  trackByFn(index: number, item: any) {
    return item.code
  }

  isSubmit: boolean = false
  visible: boolean = false
  edit: boolean = false
  filter = new CustomerFilter()
  paginationResult = new PaginationResult()
  localResult: any[] = []
  customerList: any[] = []
  marketResult: any[] = []
  marketList: any[] = []
  salesMethodResult: any[] = []
  customerTypeList: any[] = []
  loading: boolean = false
  MASTER_DATA_MANAGEMENT = MASTER_DATA_MANAGEMENT

  constructor(
    private _service: CustomerService,
    private _marketService: MarketService,
    private _customerTypeService: CustomerTypeService,
    private _localService: LocalService,
    private _salesMethodService: SalesMethodService,
    private _customerContactService: CustomerContactService,

    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
    private message: NzMessageService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách khách hàng',
        path: 'master-data/customer',
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
    this.getAllLocal()
    this.getAllSalesMethod()
    this.search()
    this.getAllMarket()
    this.getAllCustomerType()
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
    const formData = this.validateForm.getRawValue()
    this.isSubmit = false

    this._customerContactService.getall().subscribe({
      next: (contacts) => {
        this.contactList = contacts        
        if(this.edit){
          this.listOfPhone = contacts.filter(
            (item:any) => item.type === 'phone' && item.customer_Code === formData.code,
          )
          this.listOfEmail = contacts.filter(
            (item:any) => item.type === 'email' && item.customer_Code === formData.code,
          )
        }
        this._service.searchCustomer(this.filter).subscribe({
          next: (data) => {
            this.paginationResult = data
            for (let i = 0; i < data.data.length; i++) {
              const code = data.data[i].code

              // Lọc số điện thoại
              this.listOfPhoneOptions[code] = this.contactList
                .filter(
                  (item) =>
                    item.type === 'phone' &&
                    item.customer_Code === code &&
                    item.isActive == true,
                )
                .map((item) => item.value)

              // Lọc email
              this.listOfEmailOptions[code] = this.contactList
                .filter(
                  (item) =>
                    item.type === 'email' &&
                    item.customer_Code === code &&
                    item.isActive == true,
                )
                .map((item) => item.value)

              // Gán giá trị mặc định cho selectedPhones nếu chưa có
              if (this.listOfPhoneOptions[code]?.length) {
                this.selectedPhones[code] = this.listOfPhoneOptions[code][0]
              }

              // Gán giá trị mặc định cho selectedEmails nếu chưa có
              if (this.listOfEmailOptions[code]?.length) {
                this.selectedEmails[code] = this.listOfEmailOptions[code][0]
              }
            }
          },
          error: (err) => {
            console.error('Lỗi khi gọi API searchCustomer:', err)
          },
        })
      },
      error: (err) => {
        console.error('Lỗi khi gọi API getall:', err)
      },
    }) 
  }

  // getAllCustomer(){
  //   this.isSubmit = false
  //   this._service.getall().subscribe({
  //     next: (data) => {
  //       console.log(data);

  //       this.customerList = data
  //     },
  //     error: (response) => {
  //       console.log(response)
  //     },
  //   })
  // }

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

  getAllMarket() {
    this.isSubmit = false
    this._marketService.getall().subscribe({
      next: (data) => {
        this.marketList = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  getAllCustomerType() {
    this.isSubmit = false
    this._customerTypeService.getall().subscribe({
      next: (data) => {
        this.customerTypeList = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  searchMarket() {
    this.isSubmit = false
    this.marketResult = this.marketList.filter(
      (market) =>
        market.localCode === this.validateForm.get('localCode')?.value,
    )
  }

  getAllSalesMethod() {
    this.isSubmit = false
    this._salesMethodService.getall().subscribe({
      next: (data) => {
        this.salesMethodResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  exportExcel() {
    return this._service
      .exportExcelCustomer(this.filter)
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
  submitFormCustomer(): void {
    this.isSubmit = true
    // console.log(this.validateForm.getRawValue());

    if (this.validateForm.valid) {
      const formData = this.validateForm.getRawValue()
      console.log(formData)
      if (this.edit) {
        this._service.updateCustomer(formData).subscribe({
          next: (data) => {
            var contact = {
              Customer_Code: formData.code,
              Contact_List: this.listOfPhone.concat(this.listOfEmail),
            }
            this._customerContactService
              .updateCustomerContact(contact)
              .subscribe({
                next: (data) => {
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
        if (this.isCodeExist(formData.code)) {
          this.message.error(
            `Mã khu vục ${formData.code} đã tồn tại, vui lòng nhập lại`,
          )
          return
        }

        this._service.createCustomer(formData).subscribe({
          next: (data) => {
            var temp = {
              Customer_Code: data.code,
              Contact_List: this.listOfPhone.concat(this.listOfEmail),
            }
            this._customerContactService.createCustomerContact(temp).subscribe({
              next: (data) => {
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

  returnOnlyNumbers(event: KeyboardEvent) {
    const charCode = event.which ? event.which : event.keyCode
    if (charCode < 48 || charCode > 57) {
      event.preventDefault()
    }
  }
  close() {
    this.listOfPhone = []
    this.listOfEmail = []
    this.contactList.forEach((item) => {
      if (!item.isActive) {
        item.isActive = true
      }
    })
    this.visible = false
    this.resetForm()
  }

  reset() {
    this.filter = new CustomerFilter()
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
    this._service.deleteCustomer(code).subscribe({
      next: (data) => {
        this.search()
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  openEdit(data: any): void {
    this.validateForm.setValue({
      code: data.code,
      name: data.name,
      address: data.address,
      paymentTerm: data.paymentTerm,
      gap: data.gap,
      cuocVcBq: data.cuocVcBq,
      mgglhXang: data.mgglhXang,
      mgglhDau: data.mgglhDau,
      buyInfo: data.buyInfo,
      bankLoanInterest: data.bankLoanInterest,
      salesMethodCode: data.salesMethodCode,
      customerTypeCode: data.customerTypeCode,
      localCode: data.localCode,
      marketCode: data.marketCode,
      fob: data.fob,
      isActive: data.isActive,
    })
    this.listOfPhone = this.contactList.filter(
      (item) => item.type === 'phone' && item.customer_Code === data.code,
    )
    this.listOfEmail = this.contactList.filter(
      (item) => item.type === 'email' && item.customer_Code === data.code,
    )
    
    setTimeout(() => {
      this.edit = true
      this.visible = true
      this.searchMarket()
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
