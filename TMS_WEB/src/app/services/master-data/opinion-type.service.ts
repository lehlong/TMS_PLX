import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class OpinionTypeService {
  constructor(private commonService: CommonService) {}

  searchOpinionType(params: any): Observable<any> {
    return this.commonService.get('OpinionType/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('OpinionType/GetAll');
  }

  createOpinionType(params: any): Observable<any> {
    return this.commonService.post('OpinionType/Insert', params);
  }

  updateOpinionType(params: any): Observable<any> {
    return this.commonService.put('OpinionType/Update', params);
  }

  exportExcelOpinionType(params: any): Observable<any> {
    return this.commonService.downloadFile('OpinionType/Export', params);
  }

  deleteOpinionType(id: string | number): Observable<any> {
    return this.commonService.delete(`OpinionType/Delete/${id}`);
  }
}
