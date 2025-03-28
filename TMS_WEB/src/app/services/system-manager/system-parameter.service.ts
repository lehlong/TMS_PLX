import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class SystemParamaterService {
  constructor(private commonService: CommonService) {}

  search(params: any): Observable<any> {
    return this.commonService.get('XHTD/SystemParameter/Search', params);
  }

  create(params: any): Observable<any> {
    return this.commonService.post('XHTD/SystemParameter/Insert', params);
  }

  getAll(): Observable<any> {
    return this.commonService.get('XHTD/SystemParameter/GetAll');
  }
  update(params: any): Observable<any> {
    return this.commonService.put('XHTD/SystemParameter/Update', params);
  }

  exportExcel(params: any): Observable<any> {
    return this.commonService.downloadFile('XHTD/SystemParameter/Export', params);
  }

  delete(code: string | number): Observable<any> {
    return this.commonService.delete(`XHTD/SystemParameter/Delete/${code}`);
  }
}
