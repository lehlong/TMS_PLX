import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { CommonService } from '../common.service';

@Injectable({
  providedIn: 'root'
})
export class MgListTablesGroupsGroupsService {
  constructor(private commonService: CommonService) { }

  searchMgListTablesGroups(params: any): Observable<any> {
    return this.commonService.get('MgListTablesGroups/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('MgListTablesGroups/GetAll')
  }

  createMgListTablesGroups(params: any): Observable<any> {
    return this.commonService.post('MgListTablesGroups/Insert', params)
  }

  updateMgListTablesGroups(params: any): Observable<any> {
    return this.commonService.put('MgListTablesGroups/Update', params)
  }

  deleteMgListTablesGroups(timeyear: string | number): Observable<any> {
    return this.commonService.delete(`MgListTablesGroups/Delete/${timeyear}`)
  }
}
