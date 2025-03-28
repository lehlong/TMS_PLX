import { BaseFilter } from '../base.model'

export class MgOpinionListFilter extends BaseFilter {
  code: string = ''
  name: string = ''
  timeYear: number | string = ''
  auditPeriod: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
