import { Injectable } from '@angular/core'
import { CommonService } from '../common.service'
import { Observable } from 'rxjs'

@Injectable({
  providedIn: 'root',
})
export class AuditPeriodService {
  constructor(private commonService: CommonService) {}

  searchAuditPeriod(params: any): Observable<any> {
    return this.commonService.get('AuditPeriod/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('AuditPeriod/GetAll')
  }

  createAuditPeriod(params: any): Observable<any> {
    return this.commonService.post('AuditPeriod/Insert', params)
  }

  updateAuditPeriod(params: any): Observable<any> {
    return this.commonService.put('AuditPeriod/Update', params)
  }

  deleteAuditPeriod(code: string | number): Observable<any> {
    return this.commonService.delete(`AuditPeriod/Delete/${code}`)
  }
  exportExcelAuditPeriod(params: any): Observable<any> {
    return this.commonService.downloadFile('AuditPeriod/Export', params)
  }
}
