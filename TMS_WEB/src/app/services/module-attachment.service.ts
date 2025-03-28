import {Injectable} from '@angular/core';
import {CommonService} from './common.service';
import {Observable} from 'rxjs';

@Injectable({
  providedIn: 'root',
})
export class ModuleAttachmentService {
  constructor(private commonService: CommonService) {}

  Upload(file?: any, params?: any): Observable<any> {
    return this.commonService.uploadFile(`ModuleAttachment/Upload`, file, params);
  }

  GetByReferenceId(params?: any): Observable<any> {
    return this.commonService.get(`ModuleAttachment/GetByReferenceId`, params);
  }

  BatchUpload(file?: any, params?: any): Observable<any> {
    return this.commonService.uploadFiles(`ModuleAttachment/BatchUpload`, file, params);
  }
}
