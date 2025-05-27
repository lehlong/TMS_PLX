import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class ConfigSmsService {
  constructor(private commonService: CommonService) {}

  searchConfigSms(params: any): Observable<any> {
    return this.commonService.get('ConfigSms/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('ConfigSms/GetAll');
  }

  createConfigSms(params: any): Observable<any> {
    return this.commonService.post('ConfigSms/Insert', params);
  }

  updateConfigSms(params: any): Observable<any> {
    return this.commonService.put('ConfigSms/Update', params);
  }

  exportExcelConfigSms(params: any): Observable<any> {
    return this.commonService.downloadFile('ConfigSms/Export', params);
  }

  deleteConfigSms(id: string | number): Observable<any> {
    return this.commonService.delete(`ConfigSms/Delete/${id}`);
  }
}
