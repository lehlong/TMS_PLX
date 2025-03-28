import { BaseFilter } from '../base.model'

export class MgListTablesFilter extends BaseFilter {
  code: string = ''
  name: string = ''
  timeYear: number | string = ''
  auditPeriod: string = ''
  groupCode: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
