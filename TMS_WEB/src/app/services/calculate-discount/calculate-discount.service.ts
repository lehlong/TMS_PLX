import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { CommonService } from '../common.service';

@Injectable({
    providedIn: 'root',
})
export class CalculateDiscountService {
    constructor(private commonService: CommonService) { }

    search(params: any): Observable<any> {
        return this.commonService.get('CalculateDiscount/Search', params);
    }

    genarateCreate(): Observable<any> {
        return this.commonService.get('CalculateDiscount/GenarateCreate');
    }

    getByStatus(status:string): Observable<any> {
      return this.commonService.get(`CalculateDiscount/Search?keyword=${status}`);
    }
    create(input: any): Observable<any> {
        return this.commonService.post('CalculateDiscount/Create', input);
    }

    getOutput(id: any): Observable<any> {
        return this.commonService.get(`CalculateDiscount/GetOutput?id=${id}`);
    }

    getInput(id: any): Observable<any> {
        return this.commonService.get(`CalculateDiscount/GetInput?id=${id}`);
    }
    updateInput(input: any): Observable<any> {
        return this.commonService.put(`CalculateDiscount/UpdateInput`, input);
    }
    // HandleQuyTrinh(input: any): Observable<any> {
    //   return this.commonService.put(`CalculateDiscount/HandleQuyTrinh`, input);
    // }
    HandleQuyTrinh(data : any) : Observable<any>{
      return this.commonService.put(`CalculateDiscount/HandleQuyTrinh`, data)
    }
    GetHistoryAction(id: any): Observable<any> {
        return this.commonService.get(`CalculateDiscount/GetHistoryAction?code=${id}`);
    }
    exportExcel(id: any): Observable<any> {
        return this.commonService.get(`CalculateDiscount/ExportExcel?headerId=${id}`);
    }

    ExportWordTrinhKy(lstTrinhKyChecked: any, headerId : any): Observable<any> {
        return this.commonService.post(`CalculateDiscount/ExportWordTrinhKy?headerId=${headerId}`, lstTrinhKyChecked)
    }
    SendMail(model: any): Observable<any> {
      return this.commonService.get(`CalculateDiscount/SendMail?headerId=${model}`)
    }
    SendSMS(model: any): Observable<any> {
      return this.commonService.get(`CalculateDiscount/SendSMS?headerId=${model}`)
    }
    Getmail(model: any): Observable<any> {
      return this.commonService.get(`CalculateDiscount/Getmail?headerId=${model}`)
    }
    GetSms(model: any): Observable<any> {
      return this.commonService.get(`CalculateDiscount/GetSms?headerId=${model}`)
    }
    GetHistoryFile(code : any): Observable<any> {
      return this.commonService.get(`CalculateDiscount/GetHistoryFile?code=${code}`)

      }

    ExportWord(lstCustomerChecked: any[], headerId : any): Observable<any> {
        return this.commonService.post(`CalculateDiscount/ExportWord?headerId=${headerId}`, lstCustomerChecked)
    }

    GetCustomerBbdo(id : any): Observable<any> {
        return this.commonService.get(`CalculateDiscount/GetCustomerBbdo?id=${id}`)
    }

    ExportPDF(lstCustomerChecked: any, headerId: any): Observable<any> {
      return this.commonService.post(`CalculateDiscount/ExportPDF?headerId=${headerId}`, lstCustomerChecked)
    }

    GetAllInputCustomer(): Observable<any> {
      return this.commonService.get(`CalculateDiscount/GetAllInputCustomer`)
  }
}
