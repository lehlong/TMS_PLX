import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class DiscountInformationListService {
  constructor(private commonService: CommonService) {}

  searchDiscountInformationList(params: any): Observable<any> {
    return this.commonService.get('DiscountInformationList/Search', params);
  }

  getAll(): Observable<any> {
    return this.commonService.get('DiscountInformationList/GetAll');
  }

  getObjectCreate(params : any): Observable<any> {
    return this.commonService.get(`DiscountInformationList/GetObjectCreate?code=${params}`);
  }

  createData(params: any): Observable<any> {
    return this.commonService.post('DiscountInformationList/Insert', params);
  }


  createDiscountInformationList(params: any): Observable<any> {
    return this.commonService.post('DiscountInformationList/Insert', params);
  }


  updateDiscountInformationList(params: any): Observable<any> {
    return this.commonService.put('DiscountInformationList/Update', params);
  }

  exportExcelDiscountInformationList(params: any): Observable<any> {
    return this.commonService.downloadFile('DiscountInformationList/Export', params);
  }

  deleteDiscountInformationList(id: string | number): Observable<any> {
    return this.commonService.delete(`DiscountInformationList/Delete/${id}`);
  }
}
