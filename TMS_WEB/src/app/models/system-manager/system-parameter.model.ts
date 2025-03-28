import {BaseFilter} from '../base.model';

export class SystemParameterFilter extends BaseFilter {
  code: string = '';
  name: string = '';
  SortColumn?: string = '';
  IsDescending?: boolean = true;
  isActive?: boolean | string | null;
}
