import { BaseFilter } from '../base.model'
export class MapPointCustomerGoodsFilter extends BaseFilter {
  code: string = ''
  name: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
  customerCode: string = ''
  pointCode: string = ''
}
