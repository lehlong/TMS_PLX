import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class GoodsService {
  constructor(private commonService: CommonService) {}

  searchGoods(params: any): Observable<any> {
    return this.commonService.get('Goods/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('Goods/GetAll');
  }

  createGoods(params: any): Observable<any> {
    return this.commonService.post('Goods/Insert', params);
  }

  updateGoods(params: any): Observable<any> {
    return this.commonService.put('Goods/Update', params);
  }

  exportExcelGoods(params: any): Observable<any> {
    return this.commonService.downloadFile('Goods/Export', params);
  }

  deleteGoods(id: string | number): Observable<any> {
    return this.commonService.delete(`Goods/Delete/${id}`);
  }
}
