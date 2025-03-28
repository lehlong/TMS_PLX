import { Injectable } from '@angular/core'
import { CommonService } from '../common.service'
import { Observable } from 'rxjs'

@Injectable({
  providedIn: 'root',
})
export class MgListTablesService {
  constructor(private commonService: CommonService) {}

  searchMgListTables(params: any): Observable<any> {
    return this.commonService.get('MgListTables/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('MgListTables/GetAll')
  }

  createMgListTables(params: any): Observable<any> {
    return this.commonService.post('MgListTables/Insert', params)
  }

  updateMgListTables(params: any): Observable<any> {
    return this.commonService.put('MgListTables/Update', params)
  }

  deleteMgListTables(timeyear: string | number): Observable<any> {
    return this.commonService.delete(`MgListTables/Delete/${timeyear}`)
  }
  exportExcelMgListTables(params: any): Observable<any> {
    return this.commonService.downloadFile('MgListTables/Export', params)
  }
}
