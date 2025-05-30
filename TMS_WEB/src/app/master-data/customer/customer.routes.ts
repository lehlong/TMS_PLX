import { Routes } from '@angular/router'
import { CustomerDbComponent } from './customer-db/customer-db.component'
import { CustomerPtComponent } from './customer-pt/customer-pt.component'
import { CustomerTnppComponent } from './customer-tnpp/customer-tnpp.component'
import { CustomerFobComponent } from './customer-fob/customer-fob.component'
import { CustomerBbdoComponent } from './customer-bbdo/customer-bbdo.component'
import { CustomerPtsComponent } from './customer-pts/customer-pts.component'
export const customerRoutes: Routes = [
  { path: 'db', component: CustomerDbComponent },
  { path: 'pt', component: CustomerPtComponent },
  { path: 'tnpp', component: CustomerTnppComponent },
  { path: 'fob', component: CustomerFobComponent },
  { path: 'bbdo', component: CustomerBbdoComponent },
  { path: 'pts', component: CustomerPtsComponent },
]
