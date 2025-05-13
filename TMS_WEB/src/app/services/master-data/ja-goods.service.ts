import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class JaGoodsService {
  constructor(private commonService: CommonService) {}

  searchJaGoods(params: any): Observable<any> {
    return this.commonService.get('JaGoods/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('JaGoods/GetAll');
  }

  createJaGoods(params: any): Observable<any> {
    return this.commonService.post('JaGoods/Insert', params);
  }

  updateJaGoods(params: any): Observable<any> {
    return this.commonService.put('JaGoods/Update', params);
  }

  exportExcelJaGoods(params: any): Observable<any> {
    return this.commonService.downloadFile('JaGoods/Export', params);
  }

  deleteJaGoods(id: string | number): Observable<any> {
    return this.commonService.delete(`JaGoods/Delete/${id}`);
  }
}
