import AuthGuard from '../guards/auth.guard'
import { Routes } from '@angular/router'
import { UnitComponent } from './unit/unit.component'
import { LocalComponent } from './local/local.component'
import { AreaComponent } from './area/area.component'
import { AccountTypeComponent } from './account-type/account-type.component'
import { GoodsComponent } from './goods/goods.component'
import { DeliveryPointComponent } from './delivery-point/delivery-point.component'
import { CustomerTypeComponent } from './customer-type/customer-type.component'
import { WarehouseComponent } from './warehouse/warehouse.component'
import { MarketComponent } from './market/market.component'
import { TermOfPaymentComponent } from './term-of-payment/term-of-payment.component'
import { customerRoutes } from './customer/customer.routes'
import { DeliveryGroupComponent } from './delivery-group/delivery-group.component'
import { CuocVanChuyenListComponent } from './cuoc-van-chuyen-list/cuoc-van-chuyen-list.component'
import { CuocVanChuyenComponent } from './cuoc-van-chuyen/cuoc-van-chuyen.component'
import { SignerComponent } from './signer/signer.component'
import { CustomerPhoneComponent } from './customer-phone/customer-phone.component'
import { CustomerEmailComponent } from './customer-email/customer-email.component'
import { CompetitorComponent } from './competitor/competitor.component'
import { MarketCompetitorComponent } from './market-competitor/market-competitor.component'
import { JaGoodsComponent } from './ja-goods/ja-goods.component'
import { JaDiscountComponent } from './ja-discount/ja-discount.component'
import { JaPtPhanPhoiComponent } from './ja-pt-phan-phoi/ja-pt-phan-phoi.component'
import { JanaPriceComponent } from './jana-price/jana-price.component'

export const masterDataRoutes: Routes = [
  { path: 'unit', component: UnitComponent },
  { path: 'local', component: LocalComponent },
  { path: 'area', component: AreaComponent },
  { path: 'account-type', component: AccountTypeComponent },
  { path: 'delivery-group', component: DeliveryGroupComponent },
  { path: 'competitor', component: CompetitorComponent },
  { path: 'market-competitor', component: MarketCompetitorComponent },
  { path: 'goods', component: GoodsComponent },
  { path: 'customer', children: customerRoutes, canActivate: [AuthGuard]},
  { path: 'delivery-point', component: DeliveryPointComponent },
  { path: 'customer-type', component: CustomerTypeComponent },
  { path: 'warehouse', component: WarehouseComponent },
  { path: 'market', component: MarketComponent },
  { path: 'term-of-payment', component: TermOfPaymentComponent },
  { path: 'cuoc-van-chuyen/detail/:code', component: CuocVanChuyenComponent },
  { path: 'cuoc-van-chuyen-list', component: CuocVanChuyenListComponent },
  { path: 'signer', component: SignerComponent },
  { path: 'customer-phone', component: CustomerPhoneComponent },
  { path: 'customer-email', component: CustomerEmailComponent },

  { path: 'ja-goods', component: JaGoodsComponent },
  { path: 'ja-discount', component: JaDiscountComponent },
  { path: 'jana-price', component: JanaPriceComponent },
  { path: 'ja-pt-phan-phoi', component: JaPtPhanPhoiComponent },



]
