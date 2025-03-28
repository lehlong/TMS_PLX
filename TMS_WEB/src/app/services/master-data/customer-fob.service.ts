import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class CustomerFobService {
  constructor(private commonService: CommonService) {}

  searchCustomerFob(params: any): Observable<any> {
    return this.commonService.get('CustomerFob/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('CustomerFob/GetAll');
  }

  createCustomerFob(params: any): Observable<any> {
    return this.commonService.post('CustomerFob/Insert', params);
  }

  updateCustomerFob(params: any): Observable<any> {
    return this.commonService.put('CustomerFob/Update', params);
  }

  exportExcelCustomerFob(params: any): Observable<any> {
    return this.commonService.downloadFile('CustomerFob/Export', params);
  }

  deleteCustomerFob(id: string | number): Observable<any> {
    return this.commonService.delete(`CustomerFob/Delete/${id}`);
  }
}
