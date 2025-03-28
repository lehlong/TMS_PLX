import {startOfDay, endOfDay, format, subDays} from 'date-fns';

import {BaseFilter} from '../base.model';

export class ActionLogFilter extends BaseFilter {
  selectedRange: Date[] = [startOfDay(subDays(new Date(), 3)), endOfDay(new Date())];
  statusCode?: number;
}

export class ActionLogFilterRequest extends BaseFilter {
  FromDate?: string = format(subDays(new Date(), 3), 'yyyy-MM-dd');
  ToDate?: string = format(new Date(), 'yyyy-MM-dd');
}
