export enum ORDER_STATE {
  KHOI_TAO = 'KHOI_TAO',
  CHO_XAC_NHAN = 'CHO_XAC_NHAN',
  DA_XAC_NHAN = 'DA_XAC_NHAN',
  HUY = 'HUY',
  DA_IN_PHIEU = 'DA_IN_PHIEU',
  DANG_XUAT_HANG = 'DANG_XUAT_HANG',
  DA_XUAT_HANG = 'DA_XUAT_HANG',
}

export enum ORDER_ACTION {
  TAO_MOI = 'TAO_MOI',
  THAY_DOI_THONG_TIN = 'THAY_DOI_THONG_TIN',
  GUI_DON_HANG = 'GUI_DON_HANG',
  XUAT_HANG = 'XUAT_HANG',
  XAC_NHAN = 'XAC_NHAN',
  TU_CHOI = 'TU_CHOI',
  HUY = 'HUY',
  IN_PHIEU = 'IN_PHIEU',
}

export enum ORDER_TYPE {
  NPP = 'NPP',
  C2 = 'C2',
  C3 = 'C3',
}

export const OrderAction = {
  TAO_MOI: {
    name: 'Tạo mới',
    value: ORDER_ACTION.TAO_MOI,
  },
  THAY_DOI_THONG_TIN: {
    name: 'Thay đổi thông tin',
    value: ORDER_ACTION.THAY_DOI_THONG_TIN,
  },
  GUI_DON_HANG: {
    name: 'Gửi đơn hàng',
    value: ORDER_ACTION.GUI_DON_HANG,
  },
  XUAT_HANG: {
    name: 'Xuất hàng',
    value: ORDER_ACTION.XUAT_HANG,
  },
  XAC_NHAN: {
    name: 'Xác nhận',
    value: ORDER_ACTION.XAC_NHAN,
  },
  TU_CHOI: {
    name: 'Từ chối',
    value: ORDER_ACTION.TU_CHOI,
  },
  HUY: {
    name: 'Hủy',
    value: ORDER_ACTION.HUY,
  },
  IN_PHIEU: {
    name: 'In phiếu',
    value: ORDER_ACTION.IN_PHIEU,
  },
};

export const OrderState = {
  KHOI_TAO: {
    name: 'Khởi tạo',
    value: ORDER_STATE.KHOI_TAO,
    class: 'state-khoi-tao',
  },
  CHO_XAC_NHAN: {
    name: 'Chờ xác nhận',
    value: ORDER_STATE.CHO_XAC_NHAN,
    class: 'state-cho-xac-nhan',
  },
  DA_XAC_NHAN: {
    name: 'Đã xác nhận',
    value: ORDER_STATE.DA_XAC_NHAN,
    class: 'state-da-xac-nhan',
  },
  HUY: {
    name: 'Hủy',
    value: ORDER_STATE.HUY,
    class: 'state-huy',
  },
  DA_IN_PHIEU: {
    name: 'Đã in phiếu',
    value: ORDER_STATE.DA_IN_PHIEU,
    class: 'state-da-in-phieu',
  },
  DANG_XUAT_HANG: {
    name: 'Đang xuất hàng',
    value: ORDER_STATE.DANG_XUAT_HANG,
    class: 'state-dang-xuat-hang',
  },
  DA_XUAT_HANG: {
    name: 'Đã xuất hàng',
    value: ORDER_STATE.DA_XUAT_HANG,
    class: 'state-da-xuat-hang',
  },
};

export const ListOrderState = [
  {
    name: 'Khởi tạo',
    value: ORDER_STATE.KHOI_TAO,
  },
  {
    name: 'Chờ xác nhận',
    value: ORDER_STATE.CHO_XAC_NHAN,
  },
  {
    name: 'Đã xác nhận',
    value: ORDER_STATE.DA_XAC_NHAN,
  },
  {
    name: 'Hủy',
    value: ORDER_STATE.HUY,
  },
  {
    name: 'Đã in phiếu',
    value: ORDER_STATE.DA_IN_PHIEU,
  },
  {
    name: 'Đang xuất hàng',
    value: ORDER_STATE.DANG_XUAT_HANG,
  },
  {
    name: 'Đã xuất hàng',
    value: ORDER_STATE.DA_XUAT_HANG,
  },
];
