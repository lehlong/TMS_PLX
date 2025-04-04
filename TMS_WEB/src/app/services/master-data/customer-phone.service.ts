import { Injectable } from '@angular/core'
import { Observable } from 'rxjs'
import { CommonService } from '../common.service'

@Injectable({
  providedIn: 'root',
})
export class CustomerPhoneService {
  constructor(private commonService: CommonService) {}

  search(params: any): Observable<any> {
    return this.commonService.get('CustomerPhone/Search?IsActive=true', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('CustomerPhone/GetAll?IsActive=true')
  }

  create(params: any): Observable<any> {
    return this.commonService.post('CustomerPhone/Insert', params)
  }

  update(params: any): Observable<any> {
    return this.commonService.put('CustomerPhone/Update', params)
  }
}
