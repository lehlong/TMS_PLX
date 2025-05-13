import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class JaGoodsTypeService {
  constructor(private commonService: CommonService) {}

  searchJaGoodsType(params: any): Observable<any> {
    return this.commonService.get('JaGoodsType/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('JaGoodsType/GetAll');
  }

  createJaGoodsType(params: any): Observable<any> {
    return this.commonService.post('JaGoodsType/Insert', params);
  }

  updateJaGoodsType(params: any): Observable<any> {
    return this.commonService.put('JaGoodsType/Update', params);
  }

  exportExcelJaGoodsType(params: any): Observable<any> {
    return this.commonService.downloadFile('JaGoodsType/Export', params);
  }

  deleteJaGoodsType(id: string | number): Observable<any> {
    return this.commonService.delete(`JaGoodsType/Delete/${id}`);
  }
}
