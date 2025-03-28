import {Component} from '@angular/core';
import {ShareModule} from '../../../shared/share-module';
import {GlobalService} from '../../../services/global.service';
import {ProfileService} from '../../../services/system-manager/profile.service';
import {FormBuilder, FormGroup, NonNullableFormBuilder, Validators} from '@angular/forms';
@Component({
  selector: 'app-profile-index',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './profile-index.component.html',
  styleUrl: './profile-index.component.scss',
})
export class ProfileIndexComponent {
  userName: string = '';
  fullName: string = '';
  address: string = '';
  email: string = '';
  phoneNumber: string = '';

  validateForm: FormGroup = this.fb.group({
    userName: ['', [Validators.required]],
    fullName: ['', [Validators.required]],
    address: ['', [Validators.required]],
    email: ['', [Validators.required, Validators.email]],
    phoneNumber: ['', [Validators.required, Validators.pattern('^0\\d{9,10}$')]],
  });

  changePassForm: FormGroup = this._fb.group(
    {
      oldPassword: ['', [Validators.required]],
      newPassword: [
        '',
        [Validators.required, Validators.minLength(8), Validators.pattern(/^(?=.*[A-Z])(?=.*[@#$%^&+=!]).{8,}$/)],
      ],
      rePassword: [
        '',
        [Validators.required, Validators.minLength(8), Validators.pattern(/^(?=.*[A-Z])(?=.*[@#$%^&+=!]).{8,}$/)],
      ],
    },
    {validator: this.checkPasswordsMatch},
  );
  submitted: boolean = false;
  oldPassVisible = false;
  password?: string;
  newPassVisible = false;
  newPassword?: string;
  reNewPassVisible = false;
  reNewPassword?: string;
  oldPassFalse: boolean = false;
  tabVisible: number = 0;
  constructor(
    private _service: ProfileService,
    private globalService: GlobalService,
    private fb: NonNullableFormBuilder,
    private _fb: FormBuilder,
  ) {
    let UserInfo = this.globalService.getUserInfo();
    this.userName = UserInfo?.userName;
  }

  ngOnInit() {
    this.loadInforForm();
  }

  get _f() {
    return this.changePassForm.controls;
  }

  checkPasswordsMatch(formGroup: FormGroup) {
    const newPassword = formGroup.controls['newPassword'];
    const rePassword = formGroup.controls['rePassword'];
    if (!!newPassword.value && !!rePassword.value) {
      if (newPassword.value.trim() !== rePassword.value.trim()) {
        rePassword.setErrors({mismatch: true});
      } else {
        rePassword.setErrors(null);
      }
    }
  }

  changeTab(value: any) {
    this.tabVisible = value?.index;
    this.changePassForm.reset();
    this.submitted = false;
  }

  oldPassType(e: any) {
    this.oldPassFalse = false;
  }

  loadInforForm() {
    this._service.getDetail(this.userName).subscribe({
      next: (data) => {
        this.fullName = data?.fullName;
        this.address = data?.address;
        this.email = data?.email;
        this.phoneNumber = data?.phoneNumber;
        this.validateForm.setValue({
          userName: this.userName,
          fullName: data?.fullName,
          address: data?.address,
          email: data?.email,
          phoneNumber: data?.phoneNumber,
        });
      },
      error: (response) => {
        console.log(response);
      },
    });
  }

  submitForm(): void {
    if (this.validateForm.valid) {
      this._service.update(this.validateForm.getRawValue()).subscribe({
        next: (data) => {
          this.loadInforForm();
        },
        error: (response) => {
          console.log(response);
        },
      });
    } else {
      Object.values(this.validateForm.controls).forEach((control) => {
        if (control.invalid) {
          control.markAsDirty();
          control.updateValueAndValidity({onlySelf: true});
        }
      });
    }
  }

  onChangePass(): void {
    this.submitted = true;
    if (this.changePassForm.invalid) {
      return;
    }
    this._service
      .changePassWord({
        userName: this.userName,
        oldPassword: this.changePassForm.value.oldPassword.trim(),
        newPassword: this.changePassForm.value.newPassword.trim(),
      })
      .subscribe({
        next: (data: any) => {
          this.submitted = false;
          this.changePassForm?.get('oldPassword')?.setValue('');
          this.changePassForm?.get('newPassword')?.setValue('');
          this.changePassForm?.get('rePassword')?.setValue('');
        },
        error: (response) => {
          if (response.includes('MSG0104 ')) {
            this.oldPassFalse = true;
          }
        },
      });
  }

  handleClickSave(): void {
    if (this.tabVisible == 1) {
      this.submitForm();
    }
    if (this.tabVisible == 2) {
      this.onChangePass();
    }
  }
}
