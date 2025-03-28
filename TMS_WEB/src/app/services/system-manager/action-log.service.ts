import {Injectable} from '@angular/core';
import {CommonService} from '../common.service';
import {Observable} from 'rxjs';

@Injectable({
  providedIn: 'root',
})
export class ActionLogService {
  constructor(private commonService: CommonService) {}

  Search(params: any): Observable<any> {
    return this.commonService.get('ActionLog/Search', params);
  }

  Delete(id: string | number): Observable<any> {
    return this.commonService.delete(`ActionLog/Delete/${id}`);
  }
}
