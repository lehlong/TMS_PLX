import {Injectable} from '@angular/core';
import {BehaviorSubject, Observable, Subject} from 'rxjs';

@Injectable({
  providedIn: 'root',
})
export class GlobalService {
  private loading: BehaviorSubject<boolean>;
  private apiCallCount: number = 0; // Thêm bộ đếm API
  private userNameSubject: BehaviorSubject<string | null> = new BehaviorSubject<string | null>(null);

  rightSubject: Subject<string> = new Subject<string>();
  rightData: any = [];
  breadcrumbSubject: Subject<boolean> = new Subject<boolean>();
  breadcrumb: any = [];

  constructor() {
    this.loading = new BehaviorSubject<boolean>(false);
    this.rightSubject.subscribe((value) => {
      localStorage.setItem('userRights', value);
      this.rightData = value;
    });
    this.breadcrumbSubject.subscribe((value) => {
      this.breadcrumb = value;
    });
  }
  setUserName(userName: string): void {
    this.userNameSubject.next(userName);
    localStorage.setItem('userName', userName); // Lưu userName vào localStorage nếu cần
  }

  getUserName(){
    var usString : any = localStorage.getItem('userName')
   // var username = JSON.parse(usString);
    return usString;
  }

  // Phương thức để lấy userName từ localStorage khi cần
  loadUserNameFromStorage(): void {
    const storedUserName = localStorage.getItem('userName');
    if (storedUserName) {
      this.userNameSubject.next(storedUserName);
    }
  }

  setBreadcrumb(value: any) {
    localStorage.setItem('breadcrumb', JSON.stringify(value));
    this.breadcrumbSubject.next(value);
  }

  getBreadcrumb() {
    try {
      if (this.breadcrumb && this.breadcrumb?.length > 0) {
        return this.breadcrumb;
      }
      const breadcrumb = localStorage.getItem('breadcrumb');
      return breadcrumb ? JSON.parse(breadcrumb) : null;
    } catch (e) {
      return null;
    }
  }

  getUserInfo() {
    try {
      const info = localStorage.getItem('UserInfo');
      return info ? JSON.parse(info) : null;
    } catch (e) {
      return null;
    }
  }

  setUserInfo(value: any) {
    localStorage.setItem('UserInfo', JSON.stringify(value));
  }

  setRightData(data: any) {
    this.rightSubject.next(data);
    localStorage.setItem('userRights', data);
  }

  getRightData() {
    try {
      if (this.rightData?.length > 0) {
        return this.rightData;
      }
      const rights = localStorage.getItem('userRights');
      return rights ? JSON.parse(rights) : null;
    } catch (e) {
      return null;
    }
  }

  checkPermissions(permissions: string) {
    try {
      const listPermissions = this.getRightData();
      if (listPermissions) {
        return listPermissions?.includes(permissions);
      }
      return false;
    } catch (e) {
      return false;
    }
  }

  getLoading(): Observable<boolean> {
    return this.loading.asObservable();
  }

  setLoading(newValue: boolean): void {
    setTimeout(() => {
      this.loading.next(newValue);
    });
  }

  incrementApiCallCount(): void {
    this.apiCallCount++;
    this.setLoading(true);
  }

  decrementApiCallCount(): void {
    this.apiCallCount--;
    if (this.apiCallCount === 0) {
      this.setLoading(false);
    }
  }
}
