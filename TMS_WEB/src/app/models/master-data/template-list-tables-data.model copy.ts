import { BaseFilter } from '../base.model'

export class TemplateListTablesDataFilter extends BaseFilter {
  templateCode: string = ''
  orgCode: string = ''

  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
