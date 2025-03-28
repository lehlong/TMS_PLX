import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class MarketCompetitorService {
  constructor(private commonService: CommonService) {}

  searchmarket(params: any): Observable<any> {
    return this.commonService.get('MarketCompetitor/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('MarketCompetitor/GetAll');
  }

  createmarket(params: any): Observable<any> {
    return this.commonService.post('MarketCompetitor/Insert', params);
  }

  updatemarket(params: any): Observable<any> {
    return this.commonService.put('MarketCompetitor/Update', params);
  }

  exportExcelmarket(params: any): Observable<any> {
    return this.commonService.downloadFile('MarketCompetitor/Export', params);
  }

  deletemarket(id: string | number): Observable<any> {
    return this.commonService.delete(`MarketCompetitor/Delete/${id}`);
  }
}
