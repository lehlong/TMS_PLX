import { Component, Input } from '@angular/core'
import { ShareModule } from '../../../shared/share-module'
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms'
import { DropdownService } from '../../../services/dropdown/dropdown.service'
import { AccountService } from '../../../services/system-manager/account.service'
import { UserTypeCodes } from '../../../shared/constants/account.constants'
import { ActivatedRoute } from '@angular/router'
import { error } from '@ant-design/icons-angular'

@Component({
  selector: 'app-account-create',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './account-create.component.html',
  styleUrl: './account-create.component.scss',
})
export class AccountCreateComponent {
  @Input() reset: () => void = () => { }
  @Input() visible: boolean = false
  @Input() close: () => void = () => { }

  validateForm: FormGroup = this.fb.group({
    userName: ['', [Validators.required]],
    password: [''],
    fullName: ['', [Validators.required]],
    address: ['', [Validators.required]],
    phoneNumber: ['', [Validators.required, Validators.pattern("^[0-9]*$"),]],
    email: ['', [Validators.email]],
    isActive: [true],
    accountType: ['', [Validators.required]],
    organizeCode: ['', [Validators.required]],
    //partnerId: [''],
  })

  // UserTypeCodes = UserTypeCodes;
  // isShowSelectPartner: boolean = false;
  // listPartnerCustomer: any[] = [];
  passwordVisible: boolean = false
  accountType: any[] = []
  orgList: any[] = []

  constructor(
    private _service: AccountService,
    private fb: NonNullableFormBuilder,
    private dropdownService: DropdownService,
    private route: ActivatedRoute,
  ) { }

  ngOnInit(): void {
    this.loadInit()
    this.getAllAccountType()
    this.getAllOrg()
  }

  // ngAfterViewInit() {
  //   this.route.queryParams.subscribe((params) => {
  //     if (params['create_nmtv']) {
  //       this.validateForm.get('accountType')?.setValue('NM_TV')
  //     }
  //   })
  // }

  loadInit() { }

  changeSaleType(value: string) { }

  // onUserTypeChange(value: string) {
  //   if (value === 'KH') {
  //     this.isShowSelectPartner = true
  //     this.getAllPartner()
  //     const partnerIdControl = this.validateForm.get('partnerId')
  //     partnerIdControl!.setValidators([Validators.required])
  //   } else {
  //     this.isShowSelectPartner = false
  //   }
  // }

  getAllAccountType() {
    this.dropdownService
      .getAllAccountType({
        IsCustomer: true,
        SortColumn: 'name',
        IsDescending: true,
      })
      .subscribe({
        next: (data) => {
          this.accountType = data
        },
        error: (response) => {
          console.log(response)
        },
      })
  }
  getAllOrg() {
    this.dropdownService.getAllOrg().subscribe({
      next: (data) => {
        this.orgList = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  submitForm(): void {
    if (this.validateForm.valid) {
      const formValue = this.validateForm.value
      // const { partnerId, ...rest } = formValue
      // let insertObj = {}
      // if (this.isShowSelectPartner) {
      //   insertObj = formValue
      // } else {
      //   insertObj = rest
      // }

      this._service.create(formValue).subscribe({
        next: (data) => {
          this.reset()
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

  closeDrawer() {
    this.close()
    this.resetForm()
  }

  resetForm() {
    this.validateForm.reset()
  }
}
