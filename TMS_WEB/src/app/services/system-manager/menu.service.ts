import {Injectable} from '@angular/core';
import {CommonService} from '../common.service';
import {Observable} from 'rxjs';

@Injectable({
  providedIn: 'root',
})
export class MenuService {
  constructor(private commonService: CommonService) {}

  GetMenuTree() {
    return this.commonService.get('Menu/GetMenu');
  }

  GetMenuWithTreeRight(param: any) {
    return this.commonService.get('Menu/GetMenuWithTreeRight', param);
  }

  Update(data: any) {
    return this.commonService.put('Menu/Update', data);
  }

  Insert(data: any) {
    return this.commonService.post('Menu/Insert', data);
  }

  UpdateOrderTree(dataTree: any) {
    return this.commonService.put('Menu/Update-Order', dataTree);
  }

  Delete(id: string | number): Observable<any> {
    return this.commonService.delete(`Menu/Delete/${id}`);
  }
}
