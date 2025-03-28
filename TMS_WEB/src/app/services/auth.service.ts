import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { CommonService } from './common.service';

@Injectable({
  providedIn: 'root',
})
export class AuthService {
  private authEndpoints = {
    login: 'Auth/Login',
    register: 'auth/register',
    forgotPassword: 'auth/forgot-password',
    resetPassword: 'auth/reset-password',
    getRightOfUser: 'Right/getRightOfUser'
  };

  constructor(private commonService: CommonService) {}

  login(credentials: any): Observable<any> {
    return this.commonService.post(
      this.authEndpoints.login,
      credentials,
      false
    );
  }

  getRightOfUser(params: any): Observable<any> {
    return this.commonService.get(this.authEndpoints.getRightOfUser, params);
  }

  register(data: any): Observable<any> {
    return this.commonService.post(this.authEndpoints.register, data);
  }

  forgotPassword(email: string): Observable<any> {
    return this.commonService.post(this.authEndpoints.forgotPassword, {
      email,
    });
  }

  resetPassword(token: string, newPassword: string): Observable<any> {
    return this.commonService.post(this.authEndpoints.resetPassword, {
      token,
      newPassword,
    });
  }

  isLoggedIn(): boolean {
    return !!localStorage.getItem('token');
  }

  getToken(): string | null {
    return localStorage.getItem('token');
  }

  logout(): void {
    localStorage.clear();
  }
}
