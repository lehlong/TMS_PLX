import { BaseFilter } from '../base.model'
export class DeliveryPointFilter extends BaseFilter {
  code: string = ''
  name: string = ''
  customerCode: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
