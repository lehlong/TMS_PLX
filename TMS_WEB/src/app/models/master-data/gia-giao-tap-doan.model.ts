import { BaseFilter } from '../base.model'
export class GiaGiaoTapDoanFilter extends BaseFilter {
  code: string = ''
  goodsCode: string = ''
  customerCode: string = ''
  toDate: string = ''
  oldPrice: string = ''
  newPrice: string = ''

  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
