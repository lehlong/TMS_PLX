import { Injectable } from '@angular/core'
import { CommonService } from '../common.service'
import { Observable } from 'rxjs'

@Injectable({
  providedIn: 'root',
})
export class TemplateListTablesDataService {
  constructor(private commonService: CommonService) { }

  searchTemplateListTablesData(params: any): Observable<any> {
    return this.commonService.get('TemplateListTablesData/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('TemplateListTablesData/GetAll')
  }

  createTemplateListTablesData(params: any[]): Observable<any> {
    return this.commonService.post('TemplateListTablesData/Insert', params)
  }

  updateTemplateListTablesData(params: any[]): Observable<any> {
    return this.commonService.put('TemplateListTablesData/Update', params)
  }

  exportExcelDataTemplateListTablesData(templateCode: string): Observable<any> {
    return this.commonService.downloadFile(
      `TemplateListTablesData/export?templateCode=${templateCode}`,
    )
  }
  uploadExcel(file: File, auditListTablesCode: number): Observable<any> {
    const formData = new FormData()
    formData.append('file', file)
    return this.commonService.put(
      `TemplateListTablesData/UploadExcel?auditListTablesCode=${auditListTablesCode}`,
      formData,
    )
  }
  deleteTemplateListTablesData(data: {
    orgCode: string
    listTablesCode: string
    templateCode: string
  }): Observable<any> {
    return this.commonService.delete(
      `TemplateListTablesData/Delete?orgCode=${data.orgCode}&listTablesCode=${data.listTablesCode}&templateCode=${data.templateCode}`,
    )
  }
}
