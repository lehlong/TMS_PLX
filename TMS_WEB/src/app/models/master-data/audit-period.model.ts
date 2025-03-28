import { BaseFilter } from '../base.model'

export class AuditPeriodFilter extends BaseFilter {
  code: string = ''
  auditPeriod: number | string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
