import { Routes } from '@angular/router'
import { DiscountInformationComponent } from './discount-information/discount-information.component'
import { DiscountInformationListComponent } from './discount-information-list/discount-information-list.component'

export const discountInformation: Routes = [
  { path: 'detail/:code', component: DiscountInformationComponent },
  { path: 'list', component: DiscountInformationListComponent },
]
