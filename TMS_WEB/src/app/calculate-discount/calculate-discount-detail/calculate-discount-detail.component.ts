import { CustomerBbdoService } from './../../services/master-data/customer-bbdo.service';
import { Component, OnInit, output } from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { NzMessageService } from 'ng-zorro-antd/message';
import { CalculateDiscountService } from '../../services/calculate-discount/calculate-discount.service';
import { GlobalService } from '../../services/global.service';
import { ShareModule } from '../../shared/share-module';
import { environment } from '../../../environments/environment.prod'
import { isVisible } from 'ckeditor5'
import { NgxDocViewerModule } from 'ngx-doc-viewer';
import {
  CALCULATE_RESULT_RIGHT,
  IMPORT_BATCH,
} from '../../shared/constants/access-right.constants'
import { SignerService } from '../../services/master-data/signer.service';
import Swal from 'sweetalert2';
import { GoodsService } from '../../services/master-data/goods.service';

@Component({
  selector: 'app-calculate-discount-detail',
  standalone: true,
  imports: [ShareModule, NgxDocViewerModule],
  templateUrl: './calculate-discount-detail.component.html',
  styleUrl: './calculate-discount-detail.component.scss'
})
export class CalculateDiscountDetailComponent implements OnInit {
  titleTab: string = 'Dữ liệu gốc';
  loading: boolean = false;
  visibleInput: boolean = false;
  isVisibleStatus: boolean = false;
  IMPORT_BATCH = IMPORT_BATCH
  isVisiblePreview: boolean = false
  isZoom = false
  UrlOffice: string = ''
  input: any = {
    header: {},
    inputPrice: [],
    market: [],
    customerDb: [],
    customerPt: [],
    customerFob: [],
    customerTnpp: [],
    customerBbdo: [],
  };
  dataQuyTrinh: any = {
    header: {},
    status:{},
  }
  input2: any = this.input;
  rightList: any = []
  statusModel = {
    title: '',
    des: '',
    value: '',
  }

  output: any = {
    dlg: {},
    pt: [],
    db: [],
    fob: [],
    pt09: [],
    bbdo: [],
    pl1: [],
    pl2: [],
    pl3: [],
    pl4: [],
    vk11Pt: [],
    vk11Db: [],
    vk11Fob: [],
    vk11Tnpp: [],
    vk11Bb: [],
    summary: [],
  }
  checked = false
  headerId: any = '';
  signerResult: any[] = []
  isVisibleLstTrinhKy: boolean = false
  isVisibleHistory: boolean = false
  isVisibleEmail: boolean = false
  isVisibleSms: boolean = false
  isVisibleExport: boolean = false
  lstTrinhKyChecked: any[] = []
  lstSMS: any[] = []
  lstEmail: any[] = []
  lstHistoryFile: any[] = []
  lstHistory: any[] = []
  lstTrinhKy: any[] = [
    {
      code: 'CongDienKKGiaBanLe',
      name: 'Công Điện Kiểm Kê Giá Bán Lẻ',
      status: true,
    },
    { code: 'MucGiamGiaNQTM', name: 'Mức Giảm Giá NQTM', status: false },
    { code: 'QDGBanBuon', name: 'Quyết Định Giá bán Buôn', status: false },
    { code: 'QDGBanLe', name: 'Quyết Định Giá Bán lẻ', status: true },
    { code: 'QDGCtyPTS', name: 'Quyết Định Công Ty PTS', status: false },
    { code: 'QDGNoiDung', name: 'Quyết Giá Nội Dụng', status: false },
    { code: 'ToTrinh', name: 'Tờ Trình', status: false },
    { code: 'KeKhaiGia', name: 'Kê Khai Giá', status: true },
    { code: 'KeKhaiGiaChiTiet', name: 'Kê Khai Giá Chi Tiết', status: true },
  ]
  lstCustomer: any[] = []
  isVisibleCustomer: boolean = false
  isVisibleCustomerPDF: boolean = false
  lstCustomerChecked: any[] = []
  accountGroups: any = {}
  searchInput = ''
  searchTerm: { [key: string]: string } = {
    PT: "",
    DB: "",
    FOB: "",
    PT09: "",
    BBDO: "",
    PL1: "",
    PL2: "",
    PL3: "",
    PL4: "",
    VK11PT: "",
    VK11DB: "",
    VK11FOB: "",
    VK11TNPP: "",
    PTS: "",
    VK11BB: "",
    TH: ""
};
 currentTab = ""
 lstgoods: any[]= []
  constructor(
    private _service: CalculateDiscountService,
    private globalService: GlobalService,
    private message: NzMessageService,
    private route: ActivatedRoute,
    private _signerService: SignerService,
    private _goodService: GoodsService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Kết quả tính toán',
        path: 'calculate-discount/detail',
      },
    ])
    this.globalService.getLoading().subscribe((value) => {
      this.loading = value
    })
  }
  ngOnInit(): void {
    this.route.paramMap.subscribe({
      next: (params) => {
        const id = params.get('id')
        this.headerId = id
        this.getOutput(this.headerId);
      }
    })
    this._service.getInput(this.headerId).subscribe({
      next: (data) => {
        this.input = data;
        this.titleTab = data.header.name
      },
      error: (response) => {
        console.log(response)
      },
    })
    this.getRight()
  }

  getOutput(id: any) {
    this._service.getOutput(id).subscribe({
      next: (data) => {
        this.output = data;
        console.log(data)
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  getAllGood() {
    this._goodService.getall().subscribe({
      next: (data) => {
        this.lstgoods = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  showEmailAction() {
    console.log("tc")
    this._service.Getmail(this.headerId).subscribe({
      next: (data) => {
        this.lstEmail = data
        console.log(data)
        this.isVisibleEmail = true
      },
      error: (err) => {
        console.log(err)
      },
    })
  }

  removeHtmlTags(html: string): string {
    if (!html) return '';
    return html.replace(/<\/?[^>]+(>|$)/g, "");
  }

  showSMSAction() {
    this._service.GetSms(this.headerId).subscribe({
      next: (data) => {
        this.lstSMS = data
        console.log(data)
        this.isVisibleSms = true

      },
      error: (err) => {
        console.log(err)
      },
    })
  }
  confirmSendSMS() {
    console.log("err")
    this._service.SendSMS(this.headerId).subscribe({

      next: (data) => {
        this.message.create('success', 'Gửi mail thành công')
      },
      error: (err) => {
        console.log(err)
      },
    })
  }
  confirmSendsMail() {
    console.log("err")
    this._service.SendMail(this.headerId).subscribe({

      next: (data) => {
        this.message.create('success', 'Gửi mail thành công')
      },
      error: (err) => {
        console.log(err)
      },
    })
  }
  showHistoryExport() {
    this._service.GetHistoryFile(this.headerId).subscribe({
      next: (data) => {
        this.lstHistoryFile = data
        this.isVisibleExport = true
        this.lstHistoryFile.forEach((item) => {
          item.pathDownload = environment.apiUrl + item.path
          item.pathView = environment.apiUrl + item.path
        })
      },
      error: (err) => {
        console.log(err)
      },
    })
  }

  getRight() {
    const rights = localStorage.getItem('userRights');
    this.rightList = rights ? JSON.parse(rights) : [];

    const accountGroups = localStorage.getItem('UserInfo');
    this.accountGroups = accountGroups ? JSON.parse(accountGroups).accountGroups[0].name : [];
    this.accountGroups == 'G_NV_K' ? this.currentTab = 'PT':''
  }

  handleCancel() {
    this.isVisibleHistory = false
    this.isVisibleEmail = false
    this.isVisibleSms = false
    this.isVisibleExport = false
    this.isVisibleCustomer = false
    this.isVisibleCustomerPDF = false
  }
  exportWordTrinhKy() {

    this.isVisibleLstTrinhKy = !this.isVisibleLstTrinhKy
    console.log(this.isVisibleLstTrinhKy);
  }
  updateTrinhKyCheckedSet(code: any, checked: boolean): void {
    if (checked) {
      this.lstTrinhKyChecked.push(code)
    } else {
      this.lstTrinhKyChecked = this.lstTrinhKyChecked.filter((x) => x != code)
    }
  }

  onItemTrinhKyChecked(code: String, checked: boolean): void {
    this.updateTrinhKyCheckedSet(code, checked)
    console.log(this.lstTrinhKyChecked)
  }

  onAllCheckedLstTrinhKy(value: boolean): void {
    this.lstTrinhKyChecked = []
    if (value) {
      this.lstTrinhKy.forEach((i) => {
        this.lstTrinhKyChecked.push(i.code)
      })
    } else {
      this.lstTrinhKyChecked = []
    }
  }
  confirmExportWordTrinhKy() {
    if (this.lstTrinhKyChecked.length == 0) {
      this.message.create('warning', 'Vui lòng chọn trình ky xuất ra file')
      return
    } else {
      this._service
        .ExportWordTrinhKy(this.lstTrinhKyChecked, this.headerId)
        .subscribe({
          next: (data) => {
            this.lstTrinhKyChecked = []
            var a = document.createElement('a')
            a.href = environment.apiUrl + data
            a.target = '_blank'
            a.click()
            a.remove()
          },
          error: (err) => {
            console.log(err)
          },
        })
      this.lstTrinhKyChecked = []
    }
  }

  onAllChecked(value: boolean): void {
    this.lstCustomerChecked = []
    if (value) {
      this.lstCustomer.forEach((i) => {
        this.lstCustomerChecked.push({code:i.code,deliveryGroupCode:i.deliveryGroupCode})
      })
    } else {
      this.lstCustomerChecked = []
    }
  }

  confirmExportWord() {
    console.log(this.lstCustomerChecked)
    if (this.lstCustomerChecked.length == 0) {
      this.message.create(
        'warning',
        'Vui lòng chọn khách hàng cần xuất ra file',
      )
      return
    } else {
      this._service
        .ExportWord(this.lstCustomerChecked, this.headerId)
        .subscribe({
          next: (data) => {
            this.isVisibleCustomer = false
            this.lstCustomerChecked = []
            var a = document.createElement('a')
            a.href = environment.apiUrl + data
            a.target = '_blank'
            a.click()
            a.remove()
          },
          error: (err) => {
            console.log(err)
          },
        })
    }
  }

  updateCheckedSet(code: any, deliveryGroupCode: string, checked: boolean): void {
    if (checked) {
      this.lstCustomerChecked.push({ code: code, deliveryGroupCode: deliveryGroupCode });
    } else {
      this.lstCustomerChecked = this.lstCustomerChecked.filter(
        (x) => x.code !== code || x.deliveryGroupCode !== deliveryGroupCode
      );
    }
  }
  

  onItemChecked(code: String, deliveryGroupCode:string, checked: boolean): void {
    this.updateCheckedSet(code,deliveryGroupCode, checked)
  }

  onClickTab(title: string, tab: number) {
    this.titleTab = title;
  }
  changeStatus(value: string, status: string) {
    switch (value) {
      case '01':
        this.statusModel.title = 'TRÌNH DUYỆT'
        this.statusModel.des = 'Bạn có muốn Trình duyệt dữ liệu này?'
        this.input.header.status = "01"
        break
      case '02':
        this.statusModel.title = 'YÊU CẦU CHỈNH SỬA'
        this.statusModel.des = 'Bạn có muốn Yêu cầu chỉnh sửa lại dữ liệu này?'
        break
      case '03':
        this.statusModel.title = 'PHÊ DUYỆT'
        this.statusModel.des = 'Bạn có muốn Phê duyệt dữ liệu này?'
        break
      case '04':
        this.statusModel.title = 'TỪ CHỐI'
        this.statusModel.des = 'Bạn có muốn Từ chối dữ liệu này?'
        break
      case '05':
        this.statusModel.title = 'HỦY TRÌNH DUYỆT'
        this.statusModel.des = 'Bạn có muốn Hủy trình duyệt dữ liệu này?'
        break
      case '06':
        this.statusModel.title = 'HỦY PHÊ DUYỆT'
        this.statusModel.des = 'Bạn có muốn Hủy phê duyệt dữ liệu này?'
        break
    }
    this.dataQuyTrinh.status.code = status
    this.dataQuyTrinh.header = this.input.header
    this.isVisibleStatus = true
    Swal.fire({
      title: this.statusModel.title,
      text: this.statusModel.des,
      input: 'text',
      inputPlaceholder: 'Ý kiến',
      icon: 'question',
      showCancelButton: true,
      confirmButtonText: 'Đồng ý',
      cancelButtonText: 'Hủy'
    }).then((result) => {
      if (result.isConfirmed) {
        this.dataQuyTrinh.status.content = result.value

        console.log(this.dataQuyTrinh)
        this._service.HandleQuyTrinh(this.dataQuyTrinh).subscribe({
          next: (data) => {
            window.location.reload()
          }})
      }
    });

  }

  getAllSigner() {
    this._signerService.getall().subscribe({
      next: (data) => {
        this.signerResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  openInput() {
    this.getAllGood()
    this._service.getInput(this.headerId).subscribe({
      next: (data) => {
        this.input = data;
        this.input2 = structuredClone(data)
        this.formatVcfAndBvmtData()
        this.visibleInput = true;
      },
      error: (response) => {
        console.log(response)
      },
    })
    this.getAllSigner()
  }

  onUpdateInput() {
    console.log(this.input.header.date)
    this._service.updateInput(this.input).subscribe({
      next: (data) => {
      },
      error: (response) => {
        console.log(response)
      },
    })
  }


  handleQuyTrinh() {
    this._service.HandleQuyTrinh(this.input).subscribe({
      next: (data) => {
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  close(): void {
    this.visibleInput = false;
  }

  reCalculate(){
    this.getOutput(this.headerId);
  }

  fullScreen() {
    this.isZoom = true
    document.documentElement.requestFullscreen()
  }

  closeFullScreen() {
    this.isZoom = false
    document
      .exitFullscreen()
      .then(() => { })
      .catch(() => { })
  }
  showHistoryAction() {
    this._service.GetHistoryAction(this.headerId).subscribe({
      next: (data) => {
        this.lstHistory = data
        // console.log(data)
        this.isVisibleHistory = true
      },
      error: (err) => {
        console.log(err)
      },
    })
  }

  openNewTab(url: string) {
    console.log(url)
    window.open(url, '_blank')
  }

  exportExcel(){
    this._service.exportExcel(this.headerId).subscribe({
      next: (data) => {
        console.log(data)
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  Preview(url: string) {

    this.UrlOffice = url
    console.log(this.UrlOffice)
    this.isVisiblePreview = true

  }
  cancelPreview() {
    this.isVisiblePreview = !this.isVisiblePreview
  }

  onInputNumberFormat(data: any, field: string) {
    let value = data[field];
    // 1. Bỏ ký tự không hợp lệ (chỉ giữ số, '-', '.')
    value = value.replace(/[^0-9\-.]/g, '');

    // 2. Đảm bảo chỉ có 1 dấu '-' và nó đứng đầu
    const minusMatches = value.match(/-/g);
    if (minusMatches && minusMatches.length > 1) {
      value = value.replace(/-/g, ''); // Xoá hết
      value = '-' + value; // Thêm 1 dấu '-' đầu tiên
    } else if (minusMatches && !value.startsWith('-')) {
      value = value.replace(/-/g, '');
      value = '-' + value;
    }

    // 3. Xử lý dấu '.': chỉ cho sau '0' hoặc '-0' và duy nhất
    const dotIndex = value.indexOf('.');
    if (dotIndex !== -1) {
      const beforeDot = value.substring(0, dotIndex);
      const afterDot = value.substring(dotIndex + 1).replace(/\./g, '');

      if (beforeDot === '0' || beforeDot === '-0') {
        value = beforeDot + '.' + afterDot;
      } else {
        // Loại bỏ dấu '.' nếu không đúng điều kiện
        value = beforeDot + afterDot;
      }
    }

    // 4. Format phần nguyên với dấu ','
    const parts = value.split('.');
    let integerPart = parts[0].replace(/[^0-9\-]/g, ''); // giữ dấu '-'
    integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

    // 5. Ghép lại
    let formattedValue = integerPart;
    if (parts[1]) {
      formattedValue += '.' + parts[1];
    }

    // 6. Cập nhật lại giá trị hiển thị
    data[field] = formattedValue;
    // 7. Parse về số
    const rawNumber = formattedValue.replace(/,/g, '');
    const numberValue = parseFloat(rawNumber);
    const finalNumber = isNaN(numberValue) ? 0 : numberValue;
    // 8. Update vào model chuẩn
      const index = this.input2.inputPrice.findIndex((x: any) => x.goodCode === data.goodCode);
      if (index !== -1) {
        this.input.inputPrice[index][field] = finalNumber;
      }
  }

  onKeyDownNumberOnly(event: KeyboardEvent) {
    const allowedKeys = [
      'Backspace', 'ArrowLeft', 'ArrowRight', 'Delete', 'Tab', '-', '.',
    ];

    // Cho phép dùng Ctrl/Cmd kết hợp với: A, C, V, X
    if ((event.ctrlKey || event.metaKey) && ['a', 'c', 'v', 'x', 'z'].includes(event.key.toLowerCase())) {
      return;
    }

    // Cho phép nhập số hoặc phím được phép
    if ((event.key >= '0' && event.key <= '9') || allowedKeys.includes(event.key)) {
      return;
    }

    // Chặn các phím còn lại
    event.preventDefault();
  }

  formatNumber(value: any): string {
    if (value == null || value === '') return '';

    const num = parseFloat(value.toString().replace(/,/g, ''));
    if (isNaN(num)) return '';

    // Format giữ 4 chữ số sau dấu phẩy (mày có thể chỉnh lại tuỳ)
    return num.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 4 });
  }
  formatVcfAndBvmtData() {
    if (this.input2.inputPrice && Array.isArray(this.input2.inputPrice)) {
      this.input2.inputPrice.forEach((item:any) => {
        // Format các trường số cần format
        item.vcf = this.formatNumber(item.vcf);
        item.thueBvmt = this.formatNumber(item.thueBvmt);
        item.chenhLech = this.formatNumber(item.chenhLech);
        item.gblV1 = this.formatNumber(item.gblV1);
        item.gblV2 = this.formatNumber(item.gblV2);
        item.l15Blv2 = this.formatNumber(item.l15Blv2);
        item.l15Nbl = this.formatNumber(item.l15Nbl);
        item.laiGop = this.formatNumber(item.laiGop);
        item.fobV1 = this.formatNumber(item.fobV1);
        item.fobV2 = this.formatNumber(item.fobV2);
      });
    }
  }

  cancelSendSMS() { }
  cancelSendEmail() { }
  exportWord() {
    this._service.GetCustomerBbdo(this.headerId).subscribe({
      next: (data) => {
        this.lstCustomer = data
        this.isVisibleCustomer = true
      },
    })
  }
  isChecked(code: string): boolean {
    return this.lstCustomerChecked.some(item => item.code === code);
  }

    confirmExportPDF() {
    if (this.lstCustomerChecked.length == 0) {
      this.message.create(
        'warning',
        'Vui lòng chọn khách hàng cần xuất ra file',
      )
      return
    } else {
      this._service
        .ExportPDF(this.lstCustomerChecked, this.headerId)
        .subscribe({
          next: (data) => {
            this.isVisibleCustomer = false
            this.lstCustomerChecked = []
            var a = document.createElement('a')
            a.href = environment.apiUrl + data
            a.target = '_blank'
            a.click()
            a.remove()
          },
          error: (err) => {
            console.log(err)
          },
        })
    }
  }
  exportPDF() {
    this._service.GetCustomerBbdo(this.headerId).subscribe({
      next: (data) => {
        this.lstCustomer = data
        this.isVisibleCustomerPDF = true
      },
    })
  }
  search(sheetName:string) {
    this.searchTerm[sheetName] = this.searchInput;
  }

  reset(tabName:string){
    const keys = Object.keys(this.searchTerm);
    keys.forEach(key => this.searchTerm[key] = "");
    this.searchInput = "";
    this.currentTab = tabName
    console.log(tabName);
  }

  getSearchTerm(key: string): string {
    return this.searchTerm[key] || "";
  }

  hasExportPermission(): boolean {
    return this.rightList.includes(IMPORT_BATCH.EXPORT_TO_EXCEL) ||
           this.rightList.includes(IMPORT_BATCH.EXPORT_TO_PDF) ||
           this.rightList.includes(IMPORT_BATCH.EXPORT_TO_WORD);
  }
  onDateChange(date: Date) {
    
    const month = new Date(date).getMonth() + 1;
    // Tạo array mới để Angular detect thay đổi
    const temp= this.input.inputPrice.map((item1: any) => {
      const matched = this.lstgoods.find(item2 => item2.code === item1.goodCode);
      const vcfValue = matched
        ? (month >= 5 && month <= 10 ? matched.vfcHt : matched.vfcDx)
        : item1.vcf;
  
      return { ...item1, vcf: vcfValue }; // tạo object mới luôn
    });

    this.input.inputPrice = temp;
    this.input2.inputPrice = structuredClone(temp);
  
    this.formatVcfAndBvmtData();
  }
}
