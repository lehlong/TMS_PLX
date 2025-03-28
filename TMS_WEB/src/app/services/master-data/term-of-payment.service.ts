import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class TermOfPaymentService {
  constructor(private commonService: CommonService) {}

  searchTermOfPayment(params: any): Observable<any> {
    return this.commonService.get('TermOfPayment/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('TermOfPayment/GetAll');
  }

  createTermOfPayment(params: any): Observable<any> {
    return this.commonService.post('TermOfPayment/Insert', params);
  }

  updateTermOfPayment(params: any): Observable<any> {
    return this.commonService.put('TermOfPayment/Update', params);
  }

  exportExcelTermOfPayment(params: any): Observable<any> {
    return this.commonService.downloadFile('TermOfPayment/Export', params);
  }

  deleteTermOfPayment(id: string | number): Observable<any> {
    return this.commonService.delete(`TermOfPayment/Delete/${id}`);
  }
}
