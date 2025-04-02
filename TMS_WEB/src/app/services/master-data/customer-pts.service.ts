import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class CustomerPtsService {
  constructor(private commonService: CommonService) {}

  searchCustomerPts(params: any): Observable<any> {
    return this.commonService.get('CustomerPts/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('CustomerPts/GetAll');
  }

  createCustomerPts(params: any): Observable<any> {
    return this.commonService.post('CustomerPts/Insert', params);
  }

  updateCustomerPts(params: any): Observable<any> {
    return this.commonService.put('CustomerPts/Update', params);
  }

  exportExcelCustomerPts(params: any): Observable<any> {
    return this.commonService.downloadFile('CustomerPts/Export', params);
  }

  deleteCustomerPts(id: string | number): Observable<any> {
    return this.commonService.delete(`CustomerPts/Delete/${id}`);
  }
}
