import { Injectable } from '@angular/core'
import { Observable } from 'rxjs'
import { CommonService } from '../common.service'

@Injectable({
  providedIn: 'root',
})
export class CuocVanChuyenService {
  constructor(private commonService: CommonService) {}

  searchCuocVanChuyen(params: any, code: string | number): Observable<any> {
    return this.commonService.get(`CuocVanChuyen/SearchById/${code}`, params)
  }

  getall(): Observable<any> {
    return this.commonService.get('CuocVanChuyen/GetAll')
  }

  createCuocVanChuyen(params: any): Observable<any> {
    return this.commonService.post('CuocVanChuyen/importExcel', params)
  }
}
