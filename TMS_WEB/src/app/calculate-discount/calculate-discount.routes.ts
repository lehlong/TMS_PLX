import { Routes } from '@angular/router';
import { CalculateDiscountComponent } from './calculate-discount/calculate-discount.component';
import { CalculateDiscountDetailComponent } from './calculate-discount-detail/calculate-discount-detail.component';
export const calculateDiscountRoutes: Routes = [
  { path: 'list', component: CalculateDiscountComponent },
  { path: 'detail/:id', component: CalculateDiscountDetailComponent },
]
