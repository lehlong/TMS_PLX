import { BaseFilter } from '../base.model'
export class MarketCompetitorFilter extends BaseFilter {
  code: string = ''
  marketCode: string = ''
  competitorCode: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
