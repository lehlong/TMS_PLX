import {BaseFilter} from '../base.model';
export class CurrencyFilter extends BaseFilter {
  id: string = '';
  name: string = '';
  exchange_rate: number = 0;
  isActive?: boolean | string | null;
  SortColumn: string = '';
  IsDescending: boolean = true;
}
