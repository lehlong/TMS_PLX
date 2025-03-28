import {BaseFilter} from '../base.model';
export class PeriodTimeFilter extends BaseFilter {
  timeyear: number = 0;
  isClosed?: boolean | string | null;
  isDefault?: boolean | string | null;
  isActive?: boolean | string | null;
  SortColumn: string = '';
  IsDescending: boolean = true;
}
