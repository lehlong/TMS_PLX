import { BaseFilter } from '../base.model'
export class RetailPriceFilter extends BaseFilter {
  code: string = ''
  typeOfGoodsCode: string = ''
  localCode: string = ''
  toDate: string = ''
  oldPrice: string = ''
  newPrice: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
