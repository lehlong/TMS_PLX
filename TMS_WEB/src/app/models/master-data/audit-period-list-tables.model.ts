import { BaseFilter } from '../base.model'

export class AuditPeriodListTablesFilter extends BaseFilter {
  code: number | string = ''
  listTablesCode: string = ''
  auditPeriodCode: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
