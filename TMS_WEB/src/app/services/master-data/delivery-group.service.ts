import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class DeliveryGroupService {
  constructor(private commonService: CommonService) {}

  searchDeliveryGroup(params: any): Observable<any> {
    return this.commonService.get('DeliveryGroup/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('DeliveryGroup/GetAll');
  }

  createDeliveryGroup(params: any): Observable<any> {
    return this.commonService.post('DeliveryGroup/Insert', params);
  }

  updateDeliveryGroup(params: any): Observable<any> {
    return this.commonService.put('DeliveryGroup/Update', params);
  }

  exportExcelDeliveryGroup(params: any): Observable<any> {
    return this.commonService.downloadFile('DeliveryGroup/Export', params);
  }

  deleteDeliveryGroup(id: string | number): Observable<any> {
    return this.commonService.delete(`DeliveryGroup/Delete/${id}`);
  }
}
