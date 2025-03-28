import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class CustomerService {
  constructor(private commonService: CommonService) {}

  searchCustomer(params: any): Observable<any> {
    return this.commonService.get('Customer/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('Customer/GetAll');
  }

  getSpecialCustomer(): Observable<any> {
    return this.commonService.get('Customer/GetSpecialCustomer');
  }

  createCustomer(params: any): Observable<any> {
    return this.commonService.post('Customer/Insert', params);
  }

  updateCustomer(params: any): Observable<any> {
    return this.commonService.put('Customer/Update', params);
  }

  exportExcelCustomer(params: any): Observable<any> {
    return this.commonService.downloadFile('Customer/Export', params);
  }

  deleteCustomer(id: string | number): Observable<any> {
    return this.commonService.delete(`Customer/Delete/${id}`);
  }
}
