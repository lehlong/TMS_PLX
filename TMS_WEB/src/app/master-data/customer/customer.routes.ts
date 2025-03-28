
import { Routes } from '@angular/router'
import { CustomerDbComponent } from './customer-db/customer-db.component'
import { CustomerPtComponent } from './customer-pt/customer-pt.component'
export const customerRoutes: Routes = [
  { path: 'db', component: CustomerDbComponent },
  { path: 'pt', component: CustomerPtComponent },
]
