import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class ConfigTemplateService {
  constructor(private commonService: CommonService) {}

  searchConfigTemplate(params: any): Observable<any> {
    return this.commonService.get('ConfigTemplate/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('ConfigTemplate/GetAll');
  }

  createConfigTemplate(params: any): Observable<any> {
    return this.commonService.post('ConfigTemplate/Insert', params);
  }

  updateConfigTemplate(params: any): Observable<any> {
    return this.commonService.put('ConfigTemplate/Update', params);
  }

  exportExcelConfigTemplate(params: any): Observable<any> {
    return this.commonService.downloadFile('ConfigTemplate/Export', params);
  }

  deleteConfigTemplate(id: string | number): Observable<any> {
    return this.commonService.delete(`ConfigTemplate/Delete/${id}`);
  }
}
