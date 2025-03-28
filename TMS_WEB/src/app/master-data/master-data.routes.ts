import { Routes } from '@angular/router'
import { UnitComponent } from './unit/unit.component'
import { LocalComponent } from './local/local.component'
import { AreaComponent } from './area/area.component'
import { AccountTypeComponent } from './account-type/account-type.component'
import { GoodsComponent } from './goods/goods.component'
import { CustomerComponent } from './customer/customer.component'
import { DeliveryPointComponent } from './delivery-point/delivery-point.component'
import { CustomerTypeComponent } from './customer-type/customer-type.component'
import { WarehouseComponent } from './warehouse/warehouse.component'
import { MarketComponent } from './market/market.component'
import { TermOfPaymentComponent } from './term-of-payment/term-of-payment.component'
import { customerRoutes } from './customer/customer.routes'
import AuthGuard from '../guards/auth.guard'
export const masterDataRoutes: Routes = [
  { path: 'unit', component: UnitComponent },
  { path: 'local', component: LocalComponent },
  { path: 'area', component: AreaComponent },
  { path: 'account-type', component: AccountTypeComponent },

  { path: 'goods', component: GoodsComponent },
  { path: 'customer', children: customerRoutes, canActivate: [AuthGuard]},
  { path: 'delivery-point', component: DeliveryPointComponent },
  { path: 'customer-type', component: CustomerTypeComponent },
  { path: 'warehouse', component: WarehouseComponent },
  { path: 'market', component: MarketComponent },
  { path: 'term-of-payment', component: TermOfPaymentComponent },
]
