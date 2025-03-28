import { Injectable } from '@angular/core'
import { CommonService } from '../common.service'
import { Observable } from 'rxjs'

@Injectable({
  providedIn: 'root',
})
export class TemplateListTablesService {
  constructor(private commonService: CommonService) {}

  searchTemplateListTables(params: any): Observable<any> {
    return this.commonService.get('TemplateListTables/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('TemplateListTables/GetAll')
  }
  GetTemWithTreee(templateCode: string) {
    return this.commonService.get(
      `TemplateListTables/GetTemplateWithTree/${templateCode}`,
    )
  }

  createTemplateListTables(params: any): Observable<any> {
    return this.commonService.post('TemplateListTables/Insert', params)
  }

  updateTemplateListTables(params: any): Observable<any> {
    return this.commonService.put('TemplateListTables/Update', params)
  }

  exportExcelTemplateListTables(params: any): Observable<any> {
    return this.commonService.downloadFile('TemplateListTables/Export', params)
  }

  deleteTemplateListTables(id: string | number): Observable<any> {
    return this.commonService.delete(`TemplateListTables/Delete/${id}`)
  }
  ChangeIsActiveStatus(code: string): Observable<any> {
    return this.commonService.put(
      `TemplateListTables/ChangeIsActiveStatus/${code}`,
      {},
    )
  }
}
