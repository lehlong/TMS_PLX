import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { CommonService } from './common.service';

@Injectable({
  providedIn: 'root'
})
export class SidebarMenuService {
  private endpoints = {
    menuOfUser: 'Menu/getMenuOfUser',
  };

  constructor(private commonService: CommonService) {}

  getMenuOfUser(params:any): Observable<any> {
    return this.commonService.get(this.endpoints.menuOfUser, params);
  }
  getMenu(){
    return this.commonService.get('Menu/GetMenu');
  } 
}
