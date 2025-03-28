import { BaseFilter } from '../base.model'
export class CuocVanChuyenListFilter extends BaseFilter {
  code: string = ''
  name: string = ''
  localCode: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}