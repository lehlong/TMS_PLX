import { Injectable } from '@angular/core'
import { Observable } from 'rxjs'
import { CommonService } from '../common.service'

@Injectable({
  providedIn: 'root',
})
export class CustomerEmailService {
  constructor(private commonService: CommonService) {}

  search(params: any): Observable<any> {
    return this.commonService.get('CustomerEmail/Search?IsActive=true', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('CustomerEmail/GetAll?IsActive=true')
  }

  create(params: any): Observable<any> {
    return this.commonService.post('CustomerEmail/Insert', params)
  }

  update(params: any): Observable<any> {
    return this.commonService.put('CustomerEmail/Update', params)
  }
}
