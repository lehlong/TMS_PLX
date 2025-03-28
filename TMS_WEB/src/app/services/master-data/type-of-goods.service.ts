import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class TypeOfGoodsService {
  constructor(private commonService: CommonService) {}

  searchTypeOfGoods(params: any): Observable<any> {
    return this.commonService.get('TypeOfGoods/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('TypeOfGoods/GetAll');
  }

  createTypeOfGoods(params: any): Observable<any> {
    return this.commonService.post('TypeOfGoods/Insert', params);
  }

  updateTypeOfGoods(params: any): Observable<any> {
    return this.commonService.put('TypeOfGoods/Update', params);
  }

  exportExcelTypeOfGoods(params: any): Observable<any> {
    return this.commonService.downloadFile('TypeOfGoods/Export', params);
  }

  deleteTypeOfGoods(id: string | number): Observable<any> {
    return this.commonService.delete(`TypeOfGoods/Delete/${id}`);
  }
}
