import { Injectable } from '@angular/core'
import { Observable } from 'rxjs'
import { CommonService } from '../common.service'

@Injectable({
  providedIn: 'root',
})
export class CustomerContactService {
  constructor(private commonService: CommonService) {}

  createCustomerContact(params: any): Observable<any> {
    return this.commonService.post('CustomerContact/Insert', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('CustomerContact/GetAll?IsActive=true')
  }

  updateCustomerContact(params: any): Observable<any> {
    return this.commonService.put('CustomerContact/Update', params)
  }
}
