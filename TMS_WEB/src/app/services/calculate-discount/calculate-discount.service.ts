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
    copyInput(headerId:any, id: any): Observable<any> {
      return this.commonService.get(`CalculateDiscount/CopyInput?id=${id}&headerId=${headerId}`);
    }
    updateInput(input: any): Observable<any> {
        return this.commonService.put(`CalculateDiscount/UpdateInput`, input);
    }

    HandleQuyTrinh(data : any) : Observable<any>{
      return this.commonService.put(`CalculateDiscount/HandleQuyTrinh`, data)
    }
    GetHistoryAction(id: any): Observable<any> {
        return this.commonService.get(`CalculateDiscount/GetHistoryAction?code=${id}`);
    }
    exportExcel(id: any,accountGroup: any): Observable<any> {
      console.log(accountGroup)
        return this.commonService.get(`CalculateDiscount/ExportExcel?headerId=${id}&accGroup=${accountGroup}`);
    }

    ExportWordTrinhKy(lstTrinhKyChecked: any, headerId : any): Observable<any> {
        return this.commonService.postNoneMess(`CalculateDiscount/ExportWordTrinhKy?headerId=${headerId}`, lstTrinhKyChecked)
    }
  


    SaveSMS(headerId: any, smsName: any): Observable<any> {
      return this.commonService.get(`CalculateDiscount/SaveSMS?headerId=${headerId}&smsName=${smsName}`)
    }
    SendSMS(lstSms: any): Observable<any> {
      return this.commonService.post(`CalculateDiscount/SendSMS`, lstSms)
    }
    SendlstEmail(lstMail: any): Observable<any> {
      return this.commonService.post(`CalculateDiscount/SendlstMail`, lstMail)
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
        return this.commonService.postNoneMess(`CalculateDiscount/ExportWord?headerId=${headerId}`, lstCustomerChecked)
    }

    GetCustomerBbdo(id : any): Observable<any> {
        return this.commonService.get(`CalculateDiscount/GetCustomerBbdo?id=${id}`)
    }

    ExportPDF(lstCustomerChecked: any, headerId: any): Observable<any> {
      return this.commonService.postNoneMess(`CalculateDiscount/ExportPDF?headerId=${headerId}`, lstCustomerChecked)
    }

    GetAllInputCustomer(): Observable<any> {
      return this.commonService.get(`CalculateDiscount/GetAllInputCustomer`)
  }
}
