import { Routes } from '@angular/router'
import { CustomerDbComponent } from './customer-db/customer-db.component'
import { CustomerPtComponent } from './customer-pt/customer-pt.component'
import { CustomerTnppComponent } from './customer-tnpp/customer-tnpp.component'
import { CustomerFobComponent } from './customer-fob/customer-fob.component'
export const customerRoutes: Routes = [
  { path: 'db', component: CustomerDbComponent },
  { path: 'pt', component: CustomerPtComponent },
  { path: 'tnpp', component: CustomerTnppComponent },
  { path: 'fob', component: CustomerFobComponent },
]
