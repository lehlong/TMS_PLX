import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class CustomerTnppService {
  constructor(private commonService: CommonService) {}

  searchCustomerTnpp(params: any): Observable<any> {
    return this.commonService.get('CustomerTnpp/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('CustomerTnpp/GetAll');
  }

  createCustomerTnpp(params: any): Observable<any> {
    return this.commonService.post('CustomerTnpp/Insert', params);
  }

  updateCustomerTnpp(params: any): Observable<any> {
    return this.commonService.put('CustomerTnpp/Update', params);
  }

  exportExcelCustomerTnpp(params: any): Observable<any> {
    return this.commonService.downloadFile('CustomerTnpp/Export', params);
  }

  deleteCustomerTnpp(id: string | number): Observable<any> {
    return this.commonService.delete(`CustomerTnpp/Delete/${id}`);
  }
}
