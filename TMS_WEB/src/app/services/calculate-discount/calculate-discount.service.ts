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

    exportExcel(id: any): Observable<any> {
        return this.commonService.get(`CalculateDiscount/ExportExcel?headerId=${id}`);
    }
}
