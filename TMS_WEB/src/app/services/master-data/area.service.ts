import { Injectable } from '@angular/core'
import { Observable } from 'rxjs'
import { CommonService } from '../common.service'

@Injectable({
  providedIn: 'root',
})
export class AreaService {
  constructor(private commonService: CommonService) {}

  searchArea(params: any): Observable<any> {
    return this.commonService.get('Area/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('Area/GetAll')
  }

  createArea(params: any): Observable<any> {
    return this.commonService.post('Area/Insert', params)
  }

  updateArea(params: any): Observable<any> {
    return this.commonService.put('Area/Update', params)
  }

  exportExcelArea(params: any): Observable<any> {
    return this.commonService.downloadFile('Area/Export', params)
  }

  deleteArea(id: string | number): Observable<any> {
    return this.commonService.delete(`Area/Delete/${id}`)
  }
}
