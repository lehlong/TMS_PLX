import { BaseFilter } from '../base.model'
export class CompetitorFilter extends BaseFilter {
  code: string = ''
  name: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
