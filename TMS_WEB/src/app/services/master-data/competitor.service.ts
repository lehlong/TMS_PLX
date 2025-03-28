import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class CompetitorService {
  constructor(private commonService: CommonService) {}

  searchCompetitor(params: any): Observable<any> {
    return this.commonService.get('Competitor/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('Competitor/GetAll');
  }

  createCompetitor(params: any): Observable<any> {
    return this.commonService.post('Competitor/Insert', params);
  }

  updateCompetitor(params: any): Observable<any> {
    return this.commonService.put('Competitor/Update', params);
  }

  exportExcelCompetitor(params: any): Observable<any> {
    return this.commonService.downloadFile('Competitor/Export', params);
  }

  deleteCompetitor(id: string | number): Observable<any> {
    return this.commonService.delete(`Competitor/Delete/${id}`);
  }
}
