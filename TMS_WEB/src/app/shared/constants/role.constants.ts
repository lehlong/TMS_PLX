export enum EROLE_CODES {
  PHONG_KINH_DOANH = 'PHONG_KINH_DOANH',
  BAN_DIEU_HANH = 'BAN_DIEU_HANH',
  BAN_GIAM_DOC = 'BAN_GIAM_DOC',
  LAI_XE = 'LAI_XE',
}

export const RoleCodes = [
  {
    name: 'Phòng kinh doanh',
    value: EROLE_CODES.PHONG_KINH_DOANH,
  },
  {
    name: 'Ban điều hành',
    value: EROLE_CODES.BAN_DIEU_HANH,
  },
  {
    name: 'Ban giám đốc',
    value: EROLE_CODES.BAN_GIAM_DOC,
  },
  {
    name: 'Lái xe',
    value: EROLE_CODES.LAI_XE,
  },
];
