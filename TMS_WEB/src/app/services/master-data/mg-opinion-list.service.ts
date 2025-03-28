import { Injectable } from '@angular/core'
import { CommonService } from '../common.service'
import { Observable } from 'rxjs'

@Injectable({
  providedIn: 'root',
})
export class MgOpinionListService {
  constructor(private commonService: CommonService) {}

  searchMgOpinionList(params: any): Observable<any> {
    return this.commonService.get('MgOpinionList/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('MgOpinionList/GetAll')
  }

  createMgOpinionList(params: any): Observable<any> {
    return this.commonService.post('MgOpinionList/Insert', params)
  }

  updateMgOpinionList(params: any): Observable<any> {
    return this.commonService.put('MgOpinionList/Update', params)
  }

  deleteMgOpinionList(code: string | number): Observable<any> {
    return this.commonService.delete(`MgOpinionList/Delete/${code}`)
  }
  exportExcelMgOpinionList(params: any): Observable<any> {
    return this.commonService.downloadFile('MgOpinionList/Export', params)
  }
}
