export enum AUDITPERIODLISTTABLES {
  KHOI_TAO = '1',
  CHO_PHE_DUYET = '2',
  DA_PHE_DUYET = '3',
  XAC_NHAN = '4',
}

export const AuditPeriodListTablesState = [
  {
    name: 'Khởi tạo',
    value: AUDITPERIODLISTTABLES.KHOI_TAO,
  },
  {
    name: 'Chờ phê duyệt',
    value: AUDITPERIODLISTTABLES.CHO_PHE_DUYET,
  },
  {
    name: 'Đã phê duyệt',
    value: AUDITPERIODLISTTABLES.DA_PHE_DUYET,
  },
  {
    name: 'Đã xác nhận',
    value: AUDITPERIODLISTTABLES.XAC_NHAN,
  },
]
