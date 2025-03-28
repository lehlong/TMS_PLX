import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class LocalService {
  constructor(private commonService: CommonService) {}

  searchLocal(params: any): Observable<any> {
    return this.commonService.get('Local/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('Local/GetAll');
  }

  createLocal(params: any): Observable<any> {
    return this.commonService.post('Local/Insert', params);
  }

  updateLocal(params: any): Observable<any> {
    return this.commonService.put('Local/Update', params);
  }

  exportExcelLocal(params: any): Observable<any> {
    return this.commonService.downloadFile('Local/Export', params);
  }

  deleteLocal(id: string | number): Observable<any> {
    return this.commonService.delete(`Local/Delete/${id}`);
  }
}
