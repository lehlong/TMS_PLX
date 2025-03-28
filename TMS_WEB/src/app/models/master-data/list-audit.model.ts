import { BaseFilter } from '../base.model'

export enum AuditStatus {
  KHOI_TAO = 'Khoitao',
  CHO_XAC_NHAN = 'Choxacnhan',
  DA_PHE_DUYET = 'Dapheduyet',
  TU_CHOI = 'Tuchoi',
}
export class ListAuditFilter extends BaseFilter {
  code: string = ''
  name: string = ''
  timeYear: string =''
  auditPeriod: string = ''
  reportDate: Date | string | null | undefined
  reportNumber: string = ''
  status: AuditStatus = AuditStatus.KHOI_TAO
  startDate: Date | string | null | undefined
  endDate: Date | string | null | undefined
  note: string = ''
  fileId: string = ''
  opinionCode: string = ''
  isActive?: boolean | string | null
  SortColumn: string = ''
  IsDescending: boolean = true
}
export interface ListAuditInterface {
  code: string;
  name: string;
  status: AuditStatus;
}