export enum EUSERTYPE_CODES {
  NHA_MAY = 'NM',
  KHACH_HANG = 'KH',
  LAI_XE = 'LX',
  NM_THUONG_VU = 'NM_TV',
}

export const UserTypeCodes = [
  {
    name: 'Nhà máy',
    value: EUSERTYPE_CODES.NHA_MAY,
  },
  {
    name: 'Khách hàng',
    value: EUSERTYPE_CODES.KHACH_HANG,
  },
  {
    name: 'Lái xe',
    value: EUSERTYPE_CODES.LAI_XE,
  },
  {
    name: 'Nhân viên thương vụ',
    value: EUSERTYPE_CODES.NM_THUONG_VU,
  },
];
