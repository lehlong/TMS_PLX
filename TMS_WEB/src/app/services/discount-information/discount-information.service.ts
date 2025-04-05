import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class DiscountInformationService {
  constructor(private commonService: CommonService) {}

  searchDiscountInformation(params: any): Observable<any> {
    return this.commonService.get('DiscountInformation/Search', params);
  }

  getAll(code : any): Observable<any> {
    return this.commonService.get(`DiscountInformation/GetAll?code=${code}`);
  }
  UpdateDataInput(model : any): Observable<any> {
    return this.commonService.post(`DiscountInformation/UpdateDataInput`, model)
  }
  getDataInput(params : any): Observable<any> {
    return this.commonService.get(`DiscountInformation/GetDataInput?code=${params}`);
  }
  createDiscountInformation(params: any): Observable<any> {
    return this.commonService.post('DiscountInformation/Insert', params);
  }
  ExportExcel(headerId: any): Observable<any> {
    return this.commonService.get(`DiscountInformation/ExportExcel?headerId=${headerId}`)
  }
  updateDiscountInformation(params: any): Observable<any> {
    return this.commonService.put('DiscountInformation/Update', params);
  }

  exportExcelDiscountInformation(params: any): Observable<any> {
    return this.commonService.downloadFile('DiscountInformation/Export', params);
  }

  deleteDiscountInformation(id: string | number): Observable<any> {
    return this.commonService.delete(`DiscountInformation/Delete/${id}`);
  }
}
