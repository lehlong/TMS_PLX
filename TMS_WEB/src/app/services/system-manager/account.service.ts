import { Injectable } from '@angular/core'
import { Observable } from 'rxjs'
import { CommonService } from '../common.service'

@Injectable({
  providedIn: 'root',
})
export class AccountService {
  constructor(private commonService: CommonService) {}

  search(params: any): Observable<any> {
    return this.commonService.get('Account/Search', params)
  }

  getDetail(params: any): Observable<any> {
    return this.commonService.get('Account/GetDetail', params)
  }

  create(params: any): Observable<any> {
    return this.commonService.post('Account/Insert', params)
  }

  update(params: any): Observable<any> {
    return this.commonService.put('Account/Update', params)
  }

  exportExcel(params: any): Observable<any> {
    return this.commonService.downloadFile('Partner/Export', params)
  }

  delete(code: string | number): Observable<any> {
    return this.commonService.delete(`Partner/Delete/${code}`)
  }

  getByType(params: any): Observable<any> {
    return this.commonService.get('Account/GetByType', params)
  }

  exportExcelASO(params: any): Observable<any> {
    return this.commonService.downloadFile('Account/ExportASO', params)
  }
}
