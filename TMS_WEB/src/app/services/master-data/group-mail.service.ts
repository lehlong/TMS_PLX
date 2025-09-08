import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class GroupMailService {
  constructor(private commonService: CommonService) {}

  searchGroupMail(params: any): Observable<any> {
    return this.commonService.get('GroupMail/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('GroupMail/GetAll');
  }

  createGroupMail(params: any): Observable<any> {
    return this.commonService.post('GroupMail/Insert', params);
  }

  updateGroupMail(params: any): Observable<any> {
    return this.commonService.put('GroupMail/Update', params);
  }

  exportExcelGroupMail(params: any): Observable<any> {
    return this.commonService.downloadFile('GroupMail/Export', params);
  }

  deleteGroupMail(id: string | number): Observable<any> {
    return this.commonService.delete(`GroupMail/Delete/${id}`);
  }
}
