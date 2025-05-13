import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class JanaPriceService {
  constructor(private commonService: CommonService) {}

  searchJanaPrice(params: any): Observable<any> {
    return this.commonService.get('JanaPrice/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('JanaPrice/GetAll');
  }

  createJanaPrice(params: any): Observable<any> {
    return this.commonService.post('JanaPrice/Insert', params);
  }

  updateJanaPrice(params: any): Observable<any> {
    return this.commonService.put('JanaPrice/Update', params);
  }

  exportExcelJanaPrice(params: any): Observable<any> {
    return this.commonService.downloadFile('JanaPrice/Export', params);
  }

  deleteJanaPrice(id: string | number): Observable<any> {
    return this.commonService.delete(`JanaPrice/Delete/${id}`);
  }
}
