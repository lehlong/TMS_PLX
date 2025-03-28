import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class ProfileService {
  constructor(private commonService: CommonService) {}
  
  getDetail(params: any): Observable<any> {
    return this.commonService.get(`Account/GetDetail?userName=${params}`);
  }

  update(params: any): Observable<any> {
    return this.commonService.put('Account/UpdateInformation', params);
  }

  changePassWord(params: any): Observable<any> {
    return this.commonService.put('Auth/ChangePassword', params);
  }
}
