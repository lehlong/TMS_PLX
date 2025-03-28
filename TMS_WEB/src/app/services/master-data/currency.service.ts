import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class CurrencyService {
  constructor(private commonService: CommonService) {}

  searchCurrency(params: any): Observable<any> {
    return this.commonService.get('Currency/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('Currency/GetAll');
  }

  createCurrency(params: any): Observable<any> {
    return this.commonService.post('Currency/Insert', params);
  }

  updateCurrency(params: any): Observable<any> {
    return this.commonService.put('Currency/Update', params);
  }

  exportExcelCurrency(params: any): Observable<any> {
    return this.commonService.downloadFile('Currency/Export', params);
  }

  deleteCurrency(id: string | number): Observable<any> {
    return this.commonService.delete(`Currency/Delete/${id}`);
  }
}
