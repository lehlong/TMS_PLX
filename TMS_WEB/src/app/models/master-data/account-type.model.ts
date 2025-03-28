import { BaseFilter } from '../base.model'

export class AccountTypeFilter extends BaseFilter {
  id: string = ''
  name: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
