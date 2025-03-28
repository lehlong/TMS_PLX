import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class WarehouseService {
  constructor(private commonService: CommonService) {}

  searchWarehouse(params: any): Observable<any> {
    return this.commonService.get('Warehouse/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('Warehouse/GetAll');
  }

  createWarehouse(params: any): Observable<any> {
    return this.commonService.post('Warehouse/Insert', params);
  }

  updateWarehouse(params: any): Observable<any> {
    return this.commonService.put('Warehouse/Update', params);
  }

  exportExcelWarehouse(params: any): Observable<any> {
    return this.commonService.downloadFile('Warehouse/Export', params);
  }

  deleteWarehouse(id: string | number): Observable<any> {
    return this.commonService.delete(`Warehouse/Delete/${id}`);
  }
}
