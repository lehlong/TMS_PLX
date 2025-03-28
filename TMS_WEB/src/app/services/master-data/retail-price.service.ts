import { Injectable } from '@angular/core'
import { CommonService } from '../common.service'
import { Observable } from 'rxjs'

@Injectable({
  providedIn: 'root',
})
export class RetailPriceService {
  constructor(private commonService: CommonService) {}

  searchRetailPrice(params: any): Observable<any> {
    return this.commonService.get('RetailPrice/Search', params)
  }

  getall(): Observable<any> {
    return this.commonService.get('RetailPrice/GetAll')
  }
  getDataInput(): Observable<any> {
    // return this.commonService.get('GiaGiaoTapDoan/BuildDataInput')
    return this.commonService.get('RetailPrice/BuildDataInput')
  }
  createRetailPrice(params: any): Observable<any> {
    return this.commonService.post('RetailPrice/Insert', params)
  }

  updateRetailPrice(params: any): Observable<any> {
    return this.commonService.put('RetailPrice/Update', params)
  }

  deleteRetailPrice(code: string | number): Observable<any> {
    return this.commonService.delete(`RetailPrice/Delete/${code}`)
  }
  exportExcelRetailPrice(params: any): Observable<any> {
    return this.commonService.downloadFile('RetailPrice/Export', params)
  }
}
