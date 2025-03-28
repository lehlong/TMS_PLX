import {BaseFilter} from '../base.model';

export class DeviceConnectionFilter extends BaseFilter {
  id?: string;
  name?: string;
  notes?: string;
  state?: boolean;
  isActive?: boolean | string;
  createBy?: string;
  updateBy?: string;
  createDate?: string;
  updateDate?: string;

}
