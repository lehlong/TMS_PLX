import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class CustomerPtService {
  constructor(private commonService: CommonService) {}

  searchCustomerPt(params: any): Observable<any> {
    return this.commonService.get('CustomerPt/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('CustomerPt/GetAll');
  }

  createCustomerPt(params: any): Observable<any> {
    return this.commonService.post('CustomerPt/Insert', params);
  }

  updateCustomerPt(params: any): Observable<any> {
    return this.commonService.put('CustomerPt/Update', params);
  }

  exportExcelCustomerPt(params: any): Observable<any> {
    return this.commonService.downloadFile('CustomerPt/Export', params);
  }

  deleteCustomerPt(id: string | number): Observable<any> {
    return this.commonService.delete(`CustomerPt/Delete/${id}`);
  }
}
