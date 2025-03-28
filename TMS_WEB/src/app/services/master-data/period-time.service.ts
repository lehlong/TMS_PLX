import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class PeriodTimeService {
  constructor(private commonService: CommonService) {}

  searchPeriodTime(params: any): Observable<any> {
    return this.commonService.get('PeriodTime/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('PeriodTime/GetAll');
  }

  createPeriodTime(params: any): Observable<any> {
    return this.commonService.post('PeriodTime/Insert', params);
  }

  updatePeriodTime(params: any): Observable<any> {
    return this.commonService.put('PeriodTime/Update', params);
  }

  deletePeriodTime(timeyear: string | number): Observable<any> {
    return this.commonService.delete(`PeriodTime/Delete/${timeyear}`);
  }

  ChangeDefaultStatus(timeyear: string | number): Observable<any> {
    return this.commonService.put(`PeriodTime/ChangeDefault/${timeyear}`,{});
  }
  ChangeIsClosedStatus(timeyear: string | number): Observable<any> {
    return this.commonService.put(`PeriodTime/ChangeIsClosed/${timeyear}`,{});
  }
}
