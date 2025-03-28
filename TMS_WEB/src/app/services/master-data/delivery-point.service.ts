import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class DeliveryPointService {
  constructor(private commonService: CommonService) {}

  searchDeliveryPoint(params: any): Observable<any> {
    return this.commonService.get('DeliveryPoint/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('DeliveryPoint/GetAll');
  }

  createDeliveryPoint(params: any): Observable<any> {
    return this.commonService.post('DeliveryPoint/Insert', params);
  }

  updateDeliveryPoint(params: any): Observable<any> {
    return this.commonService.put('DeliveryPoint/Update', params);
  }

  exportExcelDeliveryPoint(params: any): Observable<any> {
    return this.commonService.downloadFile('DeliveryPoint/Export', params);
  }

  deleteDeliveryPoint(id: string | number): Observable<any> {
    return this.commonService.delete(`DeliveryPoint/Delete/${id}`);
  }
}
