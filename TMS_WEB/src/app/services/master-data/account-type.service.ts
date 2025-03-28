import { Injectable } from '@angular/core'
import { CommonService } from '../common.service'
import { Observable } from 'rxjs'

@Injectable({
  providedIn: 'root',
})
export class AccountTypeService {
  constructor(private commonService: CommonService) {}

  searchAccountType(params: any): Observable<any> {
    return this.commonService.get('AccountType/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('AccountType/GetAll')
  }

  createAccountType(params: any): Observable<any> {
    return this.commonService.post('AccountType/Insert', params)
  }

  updateAccountType(params: any): Observable<any> {
    return this.commonService.put('AccountType/Update', params)
  }

  deleteAccountType(id: string): Observable<any> {
    return this.commonService.delete(`AccountType/Delete/${id}`)
  }
  exportExcelAccountType(params: any): Observable<any> {
    return this.commonService.downloadFile('AccountType/Export', params)
  }
}
