import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class ConfigMailService {
  constructor(private commonService: CommonService) {}

  searchConfigMail(params: any): Observable<any> {
    return this.commonService.get('ConfigMail/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('ConfigMail/GetAll');
  }

  createConfigMail(params: any): Observable<any> {
    return this.commonService.post('ConfigMail/Insert', params);
  }

  updateConfigMail(params: any): Observable<any> {
    return this.commonService.put('ConfigMail/Update', params);
  }

  exportExcelConfigMail(params: any): Observable<any> {
    return this.commonService.downloadFile('ConfigMail/Export', params);
  }

  deleteConfigMail(id: string | number): Observable<any> {
    return this.commonService.delete(`ConfigMail/Delete/${id}`);
  }
}
