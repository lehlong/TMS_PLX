import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class CustomerBbdoService {
  constructor(private commonService: CommonService) {}

  searchCustomerBbdo(params: any): Observable<any> {
    return this.commonService.get('CustomerBbdo/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('CustomerBbdo/GetAll');
  }

  createCustomerBbdo(params: any): Observable<any> {
    return this.commonService.post('CustomerBbdo/Insert', params);
  }

  updateCustomerBbdo(params: any): Observable<any> {
    return this.commonService.put('CustomerBbdo/Update', params);
  }

  exportExcelCustomerBbdo(params: any): Observable<any> {
    return this.commonService.downloadFile('CustomerBbdo/Export', params);
  }

  deleteCustomerBbdo(id: string | number): Observable<any> {
    return this.commonService.delete(`CustomerBbdo/Delete/${id}`);
  }
}
