import { Injectable } from '@angular/core';
import { CommonService } from '../common.service';
import { Observable } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class TemplateListTablesGroupsService {
  constructor(private commonService: CommonService) { }

  searchTemplateListTablesGroups(params: any): Observable<any> {
    return this.commonService.get('TemplateListTablesGroups/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('TemplateListTablesGroups/GetAll')
  }

  createTemplateListTablesGroups(params: any): Observable<any> {
    return this.commonService.post('TemplateListTablesGroups/Insert', params)
  }

  updateTemplateListTablesGroups(params: any): Observable<any> {
    return this.commonService.put('TemplateListTablesGroups/Update', params)
  }

  deleteTemplateListTablesGroups(timeyear: string | number): Observable<any> {
    return this.commonService.delete(`TemplateListTablesGroups/Delete/${timeyear}`)
  }
}
