import { Injectable } from '@angular/core'
import { Observable } from 'rxjs'
import { CommonService } from '../common.service'

@Injectable({
  providedIn: 'root',
})
export class CuocVanChuyenListService {
  constructor(private commonService: CommonService) {}

  searchCuocVanChuyen(params: any): Observable<any> {
    return this.commonService.get('CuocVanChuyenList/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('CuocVanChuyenList/GetAll')
  }

  createCuocVanChuyenList(params: any): Observable<any> {
    return this.commonService.post('CuocVanChuyenList/Insert', params, false)
  }
}
