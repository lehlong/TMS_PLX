import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class UnitService {
  constructor(private commonService: CommonService) {}

  searchUnit(params: any): Observable<any> {
    return this.commonService.get('Unit/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('Unit/GetAll');
  }

  createUnit(params: any): Observable<any> {
    return this.commonService.post('Unit/Insert', params);
  }

  updateUnit(params: any): Observable<any> {
    return this.commonService.put('Unit/Update', params);
  }

  exportExcelUnit(params: any): Observable<any> {
    return this.commonService.downloadFile('Unit/Export', params);
  }

  deleteUnit(id: string | number): Observable<any> {
    return this.commonService.delete(`Unit/Delete/${id}`);
  }
}
