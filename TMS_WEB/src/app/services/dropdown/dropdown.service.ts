import { Injectable } from '@angular/core'
import { Observable } from 'rxjs'
import { CommonService } from '../common.service'

@Injectable({
  providedIn: 'root',
})
export class DropdownService {
  constructor(private commonService: CommonService) { }

  // đối tác
  // getAllPartner(params: any = {}, IsActive: boolean = true): Observable<any> {
  //   return this.commonService.get('Partner/GetAll', {...params, IsActive});
  // }

  // loại phương tiện
  getAllVehicleType(
    params: any = {},
    IsActive: boolean = true,
  ): Observable<any> {
    return this.commonService.get('VehicleType/GetAll', { ...params, IsActive })
  }

  // phương tiện
  getAllVehicle(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Vehicle/GetAll', { ...params, IsActive })
  }

  // huyện quận
  getAllDistrict(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('District/GetAll', { ...params, IsActive })
  }

  // phường xã
  getAllWard(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Ward/GetAll', { ...params, IsActive })
  }

  // tỉnh thành phố
  getAllProvine(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Provine/GetAll', { ...params, IsActive })
  }

  // loại hình bán hàng
  getAllSaleType(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('SaleType/GetAll', { ...params, IsActive })
  }
  // lấy tất cả tài khoản
  GetAllAccount(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Account/GetAll', { ...params, IsActive })
  }
  // Nhóm tài khoản
  getAllAccountGroup(
    params: any = {},
    IsActive: boolean = true,
  ): Observable<any> {
    return this.commonService.get('AccountGroup/GetAll', {
      ...params,
      IsActive,
    })
  }
  // lái xe
  getAllDriver(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Driver/GetAll', { ...params, IsActive })
  }

  // địa bàn
  getAllArea(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Area/GetAll', { ...params, IsActive })
  }

  // sản phẩm
  getAllItem(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Item/GetAll', { ...params, IsActive })
  }

  // đơn vị tính
  getAllUnit(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Unit/GetAll', { ...params, IsActive })
  }
  getAllCurrency(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Currency/GetAll', { ...params, IsActive })
  }

  // đơn hàng
  getAllOrder(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Order/GetAll', { ...params, IsActive })
  }

  GetAllAccountPlanVisit(
    params: any = {},
    IsActive: boolean = true,
  ): Observable<any> {
    return this.commonService.get('AccountPlanVisit/GetAll', {
      ...params,
      IsActive,
    })
  }

  // chủng loại sản phẩm
  getAllTypeProduct(
    params: any = {},
    IsActive: boolean = true,
  ): Observable<any> {
    return this.commonService.get('XHTD/TypeProduct/GetAll', {
      ...params,
      IsActive,
    })
  }
  getAllAccountType(
    params: any = {},
    IsActive: boolean = true,
  ): Observable<any> {
    return this.commonService.get('AccountType/GetAll', {
      ...params,
      IsActive,
    })
  }
  getAllPeriodTime(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('PeriodTime/GetAll', {
      ...params,
      IsActive,
    })
  }
  getAllMgListTables(
    params: any = {},
    IsActive: boolean = true,
  ): Observable<any> {
    return this.commonService.get('MgListTables/GetAll', {
      ...params,
      IsActive,
    })
  }
  getAllAudit(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('AuditPeriod/GetAll', {
      ...params,
      IsActive,
    })
  }
  getAllOrg(params: any = {}, IsActive: boolean = true): Observable<any> {
    return this.commonService.get('Organize/GetAll', {
      ...params,
      IsActive,
    })
  }
}
