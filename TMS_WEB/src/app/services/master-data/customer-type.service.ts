import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class CustomerTypeService {
  constructor(private commonService: CommonService) {}

  searchCustomerType(params: any): Observable<any> {
    return this.commonService.get('CustomerType/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('CustomerType/GetAll');
  }

  createCustomerType(params: any): Observable<any> {
    return this.commonService.post('CustomerType/Insert', params);
  }

  updateCustomerType(params: any): Observable<any> {
    return this.commonService.put('CustomerType/Update', params);
  }

  exportExcelCustomerType(params: any): Observable<any> {
    return this.commonService.downloadFile('CustomerType/Export', params);
  }

  deleteCustomerType(id: string | number): Observable<any> {
    return this.commonService.delete(`CustomerType/Delete/${id}`);
  }
}
