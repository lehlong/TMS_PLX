import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class LaiGopDieuTietService {
  constructor(private commonService: CommonService) {}

  searchLaiGopDieuTiet(params: any): Observable<any> {
    return this.commonService.get('LaiGopDieuTiet/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('LaiGopDieuTiet/GetAll');
  }

  createLaiGopDieuTiet(params: any): Observable<any> {
    return this.commonService.post('LaiGopDieuTiet/Insert', params);
  }

  updateLaiGopDieuTiet(params: any): Observable<any> {
    return this.commonService.put('LaiGopDieuTiet/Update', params);
  }

  exportExcelLaiGopDieuTiet(params: any): Observable<any> {
    return this.commonService.downloadFile('LaiGopDieuTiet/Export', params);
  }

  deleteLaiGopDieuTiet(id: string | number): Observable<any> {
    return this.commonService.delete(`LaiGopDieuTiet/Delete/${id}`);
  }
}
