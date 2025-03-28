export enum ORDER_EXECUTIVE_STATE {
  CHUA_NHAN_DON = 0,
  DA_NHAN_DON = 1,
  DA_VAO_CONG = 2,
  DA_CAN_VAO = 3,
  DANG_GOI_XE = 4,
  DANG_LAY_HANG = 5,
  DA_LAY_HANG = 6,
  DA_CAN_RA = 7,
  DA_HOAN_THANH = 8,
  DA_GIAO_HANG = 9,
  DA_XAC_THUC = 10,
}

export const ListOrderExecutiveState = [
  {
    name: 'Chưa nhận đơn',
    value: ORDER_EXECUTIVE_STATE.CHUA_NHAN_DON,
  },
  {
    name: 'Đã nhận đơn',
    value: ORDER_EXECUTIVE_STATE.DA_NHAN_DON,
  },
  {
    name: 'Đã xác thực',
    value: ORDER_EXECUTIVE_STATE.DA_XAC_THUC,
  },
  {
    name: 'Đã vào cổng',
    value: ORDER_EXECUTIVE_STATE.DA_VAO_CONG,
  },
  {
    name: 'Đã cân vào',
    value: ORDER_EXECUTIVE_STATE.DA_CAN_VAO,
  },
  {
    name: 'Đang gọi xe',
    value: ORDER_EXECUTIVE_STATE.DANG_GOI_XE,
  },
  {
    name: 'Đang lấy hàng',
    value: ORDER_EXECUTIVE_STATE.DANG_LAY_HANG,
  },
  {
    name: 'Đã lấy hàng',
    value: ORDER_EXECUTIVE_STATE.DA_LAY_HANG,
  },
  {
    name: 'Đã cân ra',
    value: ORDER_EXECUTIVE_STATE.DA_CAN_RA,
  },
  {
    name: 'Đã hoàn thành',
    value: ORDER_EXECUTIVE_STATE.DA_HOAN_THANH,
  },
  {
    name: 'Đã giao hàng',
    value: ORDER_EXECUTIVE_STATE.DA_GIAO_HANG,
  },
];

export const OrderExexutiveState = {
  0: {
    name: 'Chưa nhận đơn',
    value: ORDER_EXECUTIVE_STATE.CHUA_NHAN_DON,
    class: 'state-chua-nhan-don',
  },
  1: {
    name: 'Đã nhận đơn',
    value: ORDER_EXECUTIVE_STATE.DA_NHAN_DON,
    class: 'state-da-nhan-don',
  },
  2: {
    name: 'Đã vào cổng',
    value: ORDER_EXECUTIVE_STATE.DA_VAO_CONG,
    class: 'state-da-vao-cong',
  },
  3: {
    name: 'Đã cân vào',
    value: ORDER_EXECUTIVE_STATE.DA_CAN_VAO,
    class: 'state-da-can-vao',
  },
  4: {
    name: 'Đang gọi xe',
    value: ORDER_EXECUTIVE_STATE.DANG_GOI_XE,
    class: 'state-dang-goi-xe',
  },
  5: {
    name: 'Đang lấy hàng',
    value: ORDER_EXECUTIVE_STATE.DANG_LAY_HANG,
    class: 'state-dang-lay-hang',
  },
  6: {
    name: 'Đã lấy hàng',
    value: ORDER_EXECUTIVE_STATE.DA_LAY_HANG,
    class: 'state-da-lay-hang',
  },
  7: {
    name: 'Đã cân ra',
    value: ORDER_EXECUTIVE_STATE.DA_CAN_RA,
    class: 'state-da-can-ra',
  },
  8: {
    name: 'Đã hoàn thành',
    value: ORDER_EXECUTIVE_STATE.DA_HOAN_THANH,
    class: 'state-da-hoan-thanh',
  },
  9: {
    name: 'Đã giao hàng',
    value: ORDER_EXECUTIVE_STATE.DA_GIAO_HANG,
    class: 'state-da-giao-hang',
  },
  10: {
    name: 'Đã xác thực',
    value: ORDER_EXECUTIVE_STATE.DA_XAC_THUC,
    class: 'state-da-xac-thuc',
  },
};

export const ListAuthenticationManagementState = [
  {
    name: 'Chưa xác thực',
    value: ORDER_EXECUTIVE_STATE.DA_NHAN_DON,
  },
  {
    name: 'Đã xác thực',
    value: ORDER_EXECUTIVE_STATE.DA_XAC_THUC,
  },
  {
    name: 'Đang gọi xe',
    value: ORDER_EXECUTIVE_STATE.DANG_GOI_XE,
  },
];
