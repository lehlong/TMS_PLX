import { BaseFilter } from '../base.model'
export class CustomerFilter extends BaseFilter {
  code: string = ''
  name: string = ''
  phone: string = ''
  email: string = ''
  address: string = ''
  buyInfo: string = ''
  bankLoanInterest: string = ''
  salesMethodCode: string = ''
  localCode: string = ''
  marketCode: string = ''

  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
