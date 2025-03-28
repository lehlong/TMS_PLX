import { Injectable } from '@angular/core'
import { Observable } from 'rxjs'
import { CommonService } from '../common.service'

@Injectable({
  providedIn: 'root',
})
export class AuditPeriodListTablesService {
  constructor(private commonService: CommonService) { }
  searchAuditPeriodListTables(params: any): Observable<any> {
    return this.commonService.get('AuditPeriodListTables/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('AuditPeriodListTables/GetAll')
  }

  createAuditPeriodListTables(params: any): Observable<any> {
    return this.commonService.post('AuditPeriodListTables/Insert', params)
  }
  exportExcelListTables(params: any): Observable<any> {
    return this.commonService.downloadFile(
      'AuditPeriodListTables/Export',
      params,
    )
  }
  updateAuditPeriodListTables(params: any): Observable<any> {
    return this.commonService.put('AuditPeriodListTables/Update', params)
  }
  updateAuditTemplateListTables(params: any): Observable<any> {
    return this.commonService.put(
      'AuditPeriodListTables/UpdateMultipleAuditTemplateListTablesData',
      params,
    )
  }
  uploadExcelSTC(
    file: File,
    auditListTablesCode: number | string,
  ): Observable<any> {
    const formData = new FormData()
    formData.append('file', file)
    return this.commonService.put(
      `AuditPeriodListTables/UploadExcelSTC?auditListTablesCode=${auditListTablesCode}`,
      formData,
    )
  }
  uploadExcelDV(
    file: File,
    auditListTablesCode: number | string,
  ): Observable<any> {
    const formData = new FormData()
    formData.append('file', file)
    return this.commonService.put(
      `AuditPeriodListTables/UploadExcelDV?auditListTablesCode=${auditListTablesCode}`,
      formData,
    )
  }
  ChangeStatusReview(
    code: string | number,
    textContent: string,
  ): Observable<any> {
    return this.commonService.put(
      `AuditPeriodListTables/ChangeStatusReview/${code}?textContent=${encodeURIComponent(textContent)}`,
      {},
    )
  }

  ChangeStatusCancel(
    code: string | number,
    action: string,
    textContent: string,
  ): Observable<any> {
    return this.commonService.put(
      `AuditPeriodListTables/ChangeStatusCancel/${code}?action=${encodeURIComponent(action)}&textContent=${encodeURIComponent(textContent)}`,
      {},
    )
  }
  ChangeStatusConfirm(
    code: string | number,
    action: string,
    textContent: string,
  ): Observable<any> {
    return this.commonService.put(
      `AuditPeriodListTables/ChangeStatusConfirm/${code}?action=${encodeURIComponent(action)}&textContent=${encodeURIComponent(textContent)}`,
      {},
    )
  }

  ChangeStatusApproval(code: string | number): Observable<any> {
    return this.commonService.put(
      `AuditPeriodListTables/ChangeStatusApproval/${code}`,
      {},
    )
  }

  deleteAuditPeriodListTables(code: string | number): Observable<any> {
    return this.commonService.delete(`AuditPeriodListTables/Delete/${code}`)
  }
  GetMgListTablesByAuditPeriodCode(auditPeriodCode: string): Observable<any> {
    return this.commonService.get(
      `AuditPeriodListTables/GetMgListTablesByAuditPeriodCode/${auditPeriodCode}`,
    )
  }
  GetTemplateListTablesByAuditPeriodCode(params: any): Observable<any> {
    return this.commonService.get(
      'AuditPeriodListTables/GetTemplateListTablesByAuditPeriodCode', params
    )
  }
  GetHistoryByListAuditCode(listAuditCode: string): Observable<any> {
    return this.commonService.get(
      `AuditPeriodListTables/GetHistoryByListAuditCode/${listAuditCode}`,
    )
  }
  GetTemDataWithMgListTablesAndOrgCode(params: any): Observable<any> {
    return this.commonService.get(
      'AuditPeriodListTables/GetTemDataWithMgListTablesAndOrgCode',
      params,
    )
  }
}
