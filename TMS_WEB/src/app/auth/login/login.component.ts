import { Component } from '@angular/core'
import {
  FormControl,
  FormGroup,
  NonNullableFormBuilder,
  Validators,
} from '@angular/forms'
import { ShareModule } from '../../shared/share-module/index'
import { AuthService } from '../../services/auth.service'
import { Router } from '@angular/router'
import { GlobalService } from '../../services/global.service'

@Component({
  selector: 'app-login',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './login.component.html',
  styleUrl: './login.component.scss',
})
export class LoginComponent {
  validateForm: FormGroup<{
    userName: FormControl<string>
    password: FormControl<string>
    remember: FormControl<boolean>
  }> = this.fb.group({
    userName: ['', [Validators.required]],
    password: ['', [Validators.required]],
    remember: [true],
  })

  constructor(
    private fb: NonNullableFormBuilder,
    private authService: AuthService,
    private router: Router,
    private globalService: GlobalService,
  ) {}
  passwordVisible: boolean = false
  submitForm(): void {
    if (this.validateForm.valid) {
      try {
        const credentials = {
          userName: this.validateForm.get('userName')?.value,
          password: this.validateForm.get('password')?.value,
        }

        this.authService.login(credentials).subscribe({
          next: (response) => {
            localStorage.setItem('token', response.accessToken)
            localStorage.setItem('refreshToken', response.refreshToken)
            this.globalService.setUserInfo(response.accountInfo)
            localStorage.setItem('openSidebar', 'true')
            const userName = response?.accountInfo?.userName;
            if (userName) {
              this.globalService.setUserName(userName); // Lưu userName
            }
            this.authService
              .getRightOfUser({ userName: response?.accountInfo?.userName })
              .subscribe({
                next: (rights) => {
                  this.globalService.setRightData(JSON.stringify(rights || []))
                  this.router.navigate(['/'])
                },
                error: (error) => {
                  console.error('Get right of user failed:', error)
                },
              })
          },
          error: (error) => {
            console.error('Login failed:', error)
          },
        })
      } catch (e) {
        console.error('Login failed:', e)
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
}
