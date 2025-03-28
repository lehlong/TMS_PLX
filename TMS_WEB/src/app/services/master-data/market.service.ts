import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class MarketService {
  constructor(private commonService: CommonService) {}

  searchmarket(params: any): Observable<any> {
    return this.commonService.get('Market/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('Market/GetAll');
  }

  createmarket(params: any): Observable<any> {
    return this.commonService.post('Market/Insert', params);
  }

  updatemarket(params: any): Observable<any> {
    return this.commonService.put('Market/Update', params);
  }

  exportExcelmarket(params: any): Observable<any> {
    return this.commonService.downloadFile('Market/Export', params);
  }

  deletemarket(id: string | number): Observable<any> {
    return this.commonService.delete(`Market/Delete/${id}`);
  }
}
