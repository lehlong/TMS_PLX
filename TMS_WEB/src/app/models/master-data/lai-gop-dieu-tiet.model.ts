import { BaseFilter } from '../base.model'
export class LaiGopDieuTietFilter extends BaseFilter {
  code: string = ''
  goodsCode: string = ''
  marketCode: string = ''
  price: string = ''
  toDate: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
