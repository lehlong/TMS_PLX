import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class ListAuditService {
  constructor(private commonService: CommonService) {}

  searchListAudit(params: any): Observable<any> {
    return this.commonService.get('ListAudit/Search', params);
  }
  getListAuditByCode(code: string): Observable<any> {
    return this.commonService.get(`ListAudit/GetByCode/${code}`);
  }
  getListAuditHistory(code: string): Observable<any> {
    return this.commonService.get(`ListAudit/GetListAuditHistory/${code}`);
  }
  getListOrg(code: string): Observable<any> {
    return this.commonService.get(`ListAudit/GetListOrg/${code}`);
  }
  getListOpinionUnfinished(TimeYear: string, AuditPeriod: string): Observable<any>{
    return this.commonService.get('ListAudit/GetOpinionListUnfinished', {TimeYear, AuditPeriod})
  }
  getOpinionStatis(code : string): Observable<any> {
    return this.commonService.get(`ListAudit/getOpinionStatis/${code}`);
  }

  getall(): Observable<any> {
    return this.commonService.get('ListAudit/GetAll');
  }

  createListAudit(params: any): Observable<any> {
    return this.commonService.post('ListAudit/Insert', params);
  }

  updateListAudit(params: any): Observable<any> {
    return this.commonService.put('ListAudit/Update', params);
  }

  UpdateAudit(params: any): Observable<any> {
    return this.commonService.put('ListAudit/UpdateListAudit', params);
  }

  exportExcelListAudit(params: any): Observable<any> {
    return this.commonService.downloadFile('ListAudit/Export', params);
  }

  deleteListAudit(id: string | number): Observable<any> {
    return this.commonService.delete(`ListAudit/Delete/${id}`);
  }
  

  getListFile(referenceId: string): Observable<any> {
    const params = { ReferenceId: referenceId };
    return this.commonService.get('ModuleAttachment/GetByReferenceId', params);
  }
}
