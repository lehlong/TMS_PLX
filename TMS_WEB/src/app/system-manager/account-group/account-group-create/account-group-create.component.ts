import { Component, Input } from '@angular/core'
import { ShareModule } from '../../../shared/share-module'
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms'
import { AccountGroupService } from '../../../services/system-manager/account-group.service'
import { ACCOUNT_GROUP_RIGHTS } from '../../../shared/constants/index'
@Component({
  selector: 'app-account-group-create',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './account-group-create.component.html',
  styleUrls: ['./account-group-create.component.scss'],
})
export class AccountGroupCreateComponent {
  @Input() reset: () => void = () => { }
  @Input() visible: boolean = false
  @Input() loading: boolean = false
  @Input() close: () => void = () => { }
  ACCOUNT_GROUP_RIGHTS = ACCOUNT_GROUP_RIGHTS
  validateForm: FormGroup = this.fb.group({
    name: ['', [Validators.required]],
    notes: [''],
    isActive: [true, [Validators.required]],
  })

  constructor(
    private _service: AccountGroupService,
    private fb: NonNullableFormBuilder,
  ) { }

  ngAfterViewInit() {
    this.resetForm()
  }
  ngOnInit(): void {
    this.loadInit()
  }
  loadInit() { }

  submitForm(): void {
    if (this.validateForm.valid) {
      this._service.Insert(this.validateForm.value).subscribe({
        next: (data) => {
          this.reset()
        },
        error: (response) => { },
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
