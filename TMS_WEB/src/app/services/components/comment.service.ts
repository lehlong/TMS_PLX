import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';
@Injectable({
  providedIn: 'root',
})
export class CommentService {
  constructor(private commonService: CommonService) {}

  getAllByReference(params: any): Observable<any> {
    return this.commonService.get('Comment/GetAllByReference', params);
  }

  create(params: any): Observable<any> {
    return this.commonService.post('Comment/Insert', params);
  }

  update(params: any): Observable<any> {
    return this.commonService.put('Comment/Update', params);
  }

  delete(id: string | number): Observable<any> {
    return this.commonService.delete(`Comment/Delete?Id=${id}`);
  }
}
