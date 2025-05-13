import {Injectable} from '@angular/core';
import {Observable} from 'rxjs';
import {CommonService} from '../common.service';

@Injectable({
  providedIn: 'root',
})
export class JaDiscountService {
  constructor(private commonService: CommonService) {}

  searchJaDiscount(params: any): Observable<any> {
    return this.commonService.get('JaDiscount/Search', params);
  }

  getall(): Observable<any> {
    return this.commonService.get('JaDiscount/GetAll');
  }

  createJaDiscount(params: any): Observable<any> {
    return this.commonService.post('JaDiscount/Insert', params);
  }

  updateJaDiscount(params: any): Observable<any> {
    return this.commonService.put('JaDiscount/Update', params);
  }

  exportExcelJaDiscount(params: any): Observable<any> {
    return this.commonService.downloadFile('JaDiscount/Export', params);
  }

  deleteJaDiscount(id: string | number): Observable<any> {
    return this.commonService.delete(`JaDiscount/Delete/${id}`);
  }
}
