import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class CustomerDbService {
  constructor(private commonService: CommonService) {}

  searchCustomerDb(params: any): Observable<any> {
    return this.commonService.get('CustomerDb/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('CustomerDb/GetAll');
  }

  createCustomerDb(params: any): Observable<any> {
    return this.commonService.post('CustomerDb/Insert', params);
  }

  updateCustomerDb(params: any): Observable<any> {
    return this.commonService.put('CustomerDb/Update', params);
  }

  exportExcelCustomerDb(params: any): Observable<any> {
    return this.commonService.downloadFile('CustomerDb/Export', params);
  }

  deleteCustomerDb(id: string | number): Observable<any> {
    return this.commonService.delete(`CustomerDb/Delete/${id}`);
  }
}
