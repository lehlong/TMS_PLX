import {Injectable} from '@angular/core';
import {CommonService} from '../common.service';
import {Observable} from 'rxjs';
import {AccountGroupFilter} from '../../models/system-manager/account-group.model';

@Injectable({
  providedIn: 'root',
})
export class AccountGroupService {
  constructor(private _commonService: CommonService) {}

  search(params: any): Observable<any> {
    return this._commonService.get(`AccountGroup/Search`, params);
  }

  // GetDetail(id: string | number) {
  //   return this._commonService.get(`AccountGroup/GetDetail?code=${id}`);
  // }

  GetDetail(id: string | number): Observable<any> {
    return this._commonService.get(`AccountGroup/GetDetail?code=${id}`);
  }

  Insert(params: any): Observable<any> {
    return this._commonService.post(`AccountGroup/Insert`, params);
  }

  Update(params: any): Observable<any> {
    return this._commonService.put(`AccountGroup/Update`, params);
  }

  delete(id: string | number): Observable<any> {
    return this._commonService.delete(`AccountGroup/Delete/${id}`);
  }
  ExportExcel(params: any): Observable<any> {
    return this._commonService.downloadFile(`AccountGroup/Export`, params);
  }
}
