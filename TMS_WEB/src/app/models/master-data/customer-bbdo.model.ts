import { BaseFilter } from '../base.model'
export class CustomerBbdoFilter extends BaseFilter {
  code: string = ''
  name: string = ''
  createDate: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
