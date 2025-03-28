import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class MapPointCustomerGoodsService {
  constructor(private commonService: CommonService) {}

  searchMapPointCustomerGoods(params: any): Observable<any> {
    return this.commonService.get('MapPointCustomerGoods/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('MapPointCustomerGoods/GetAll');
  }

  createMapPointCustomerGoods(params: any): Observable<any> {
    return this.commonService.post('MapPointCustomerGoods/Insert', params);
  }

  updateMapPointCustomerGoods(params: any): Observable<any> {
    return this.commonService.put('MapPointCustomerGoods/Update', params);
  }

  exportExcelMapPointCustomerGoods(params: any): Observable<any> {
    return this.commonService.downloadFile('MapPointCustomerGoods/Export', params);
  }

  deleteMapPointCustomerGoods(id: string | number): Observable<any> {
    return this.commonService.delete(`MapPointCustomerGoods/Delete/${id}`);
  }
}
