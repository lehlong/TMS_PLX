import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class JaPtPhanPhoiService {
  constructor(private commonService: CommonService) {}

  searchJaPtPhanPhoi(params: any): Observable<any> {
    return this.commonService.get('JaPtPhanPhoi/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('JaPtPhanPhoi/GetAll');
  }

  createJaPtPhanPhoi(params: any): Observable<any> {
    return this.commonService.post('JaPtPhanPhoi/Insert', params);
  }

  updateJaPtPhanPhoi(params: any): Observable<any> {
    return this.commonService.put('JaPtPhanPhoi/Update', params);
  }

  exportExcelJaPtPhanPhoi(params: any): Observable<any> {
    return this.commonService.downloadFile('JaPtPhanPhoi/Export', params);
  }

  deleteJaPtPhanPhoi(id: string | number): Observable<any> {
    return this.commonService.delete(`JaPtPhanPhoi/Delete/${id}`);
  }
}
