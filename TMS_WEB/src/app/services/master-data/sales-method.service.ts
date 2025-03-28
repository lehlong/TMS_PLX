import { Injectable } from '@angular/core'
import { CommonService } from '../common.service'
import { Observable } from 'rxjs'

@Injectable({
  providedIn: 'root',
})
export class SalesMethodService {
  constructor(private commonService: CommonService) {}

  searchSalesMethod(params: any): Observable<any> {
    return this.commonService.get('SalesMethod/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('SalesMethod/GetAll')
  }

  createSalesMethod(params: any): Observable<any> {
    return this.commonService.post('SalesMethod/Insert', params)
  }

  updateSalesMethod(params: any): Observable<any> {
    return this.commonService.put('SalesMethod/Update', params)
  }

  deleteSalesMethod(code: string | number): Observable<any> {
    return this.commonService.delete(`SalesMethod/Delete/${code}`)
  }
  exportExcelSalesMethod(params: any): Observable<any> {
    return this.commonService.downloadFile('SalesMethod/Export', params)
  }
}
