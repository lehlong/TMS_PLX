import { BaseFilter } from '../base.model'
export class CustomerTypeFilter extends BaseFilter {
  code: string = ''
  name: string = ''
  thueBvmt: string = ''
  createDate: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
