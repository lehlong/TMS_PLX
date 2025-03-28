export enum SALE_TYPE {
  NPP = 'NPP',
  C2 = 'CH-C2',
  C3 = 'CH-C3',
}

export const saleType = {
  NPP: {
    name: 'Nhà phân phối',
    value: SALE_TYPE.NPP,
  },
  C2: {
    name: 'Cửa hàng C2',
    value: SALE_TYPE.C2,
  },
  C3: {
    name: 'Cửa hàng C3',
    value: SALE_TYPE.C3,
  },
}
