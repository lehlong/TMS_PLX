import { CustomerBbdoService } from './../../services/master-data/customer-bbdo.service'
import { Component, OnInit, output } from '@angular/core'
import { ActivatedRoute, Router } from '@angular/router'

import { CalculateDiscountService } from '../../services/calculate-discount/calculate-discount.service'
import { GlobalService } from '../../services/global.service'
import { ShareModule } from '../../shared/share-module'
import { environment } from '../../../environments/environment.prod'
import { isVisible } from 'ckeditor5'
import { NgxDocViewerModule } from 'ngx-doc-viewer'
import {
  CALCULATE_RESULT_RIGHT,
  IMPORT_BATCH,
} from '../../shared/constants/access-right.constants'
import { SignerService } from '../../services/master-data/signer.service'
import Swal from 'sweetalert2'
import { GoodsService } from '../../services/master-data/goods.service'
import { DocumentEditorModule } from '@onlyoffice/document-editor-angular';
import { IConfig } from '@onlyoffice/document-editor-angular';
import { isPlatformBrowser, CommonModule } from '@angular/common';
import { NzMessageService } from 'ng-zorro-antd/message'
import { iif } from 'rxjs'
import { ConfigTemplateService } from '../../services/system-manager/config-template.service'
import { Route } from '@angular/router';
import { Location } from '@angular/common';
import { LocalService } from '../../services/master-data/local.service'

@Component({
  selector: 'app-calculate-discount-detail',
  standalone: true,
  imports: [ShareModule, NgxDocViewerModule, DocumentEditorModule],
  templateUrl: './calculate-discount-detail.component.html',
  styleUrl: './calculate-discount-detail.component.scss',
})
export class CalculateDiscountDetailComponent implements OnInit {
  titleTab: string = 'Dữ liệu gốc'
  loading: boolean = false
  visibleInput: boolean = false
  isVisibleStatus: boolean = false
  IMPORT_BATCH = IMPORT_BATCH
  isVisiblePreview: boolean = false
  isZoom = false
  UrlOffice: string = ''
  inputSearchCustomer: string = ''
  inputSearchMail: string = ''
  inputnameBBDO: string = ''
  listNameBBDO: any[] = []

  input: any = {
    header: {},
    inputPrice: [],
    market: [],
    customerDb: [],
    customerPt: [],
    customerFob: [],
    customerTnpp: [],
    customerBbdo: [],
  }
  dataQuyTrinh: any = {
    header: {},
    status: {},
  }
  input2: any = this.input
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
  visibleColSearch: boolean = false
  headerId: any = ''
  signerResult: any[] = []
  isVisibleLstTrinhKy: boolean = false
  isSms: boolean = false
  isVisibleHistory: boolean = false
  isVisibleEmail: boolean = false
  isVisibleSms: boolean = false
  isVisibleExport: boolean = false
  lstTrinhKyChecked: any[] = []
  lstSMS: any[] = []
  lstEmail: any[] = []
  lstHistoryFile: any[] = []
  lstHistory: any[] = []
  lstSms: any[] = []
  searchValue = '';
  visible = false;
  listOfData: any[] = []
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
    { code: 'ThueBvmtPts', name: 'Thuế BVMT BTS', status: true },
  ]
  lstCustomer: any[] = []
  isVisibleCustomer: boolean = false
  isVisibleCustomerPDF: boolean = false
  lstCustomerChecked: any[] = []
  lstSendSmsChecked: any[] = []
  lstSendEmailChecked: any[] = []
  EmailName = ''
  accountGroups: any = {}
  searchInput = ''
  searchInputTab = ''
  isConfirmLoading = false;
  searchTerm: { [key: string]: string } = {
    PT: '',
    DB: '',
    FOB: '',
    PT09: '',
    BBDO: '',
    PL1: '',
    PL2: '',
    PL3: '',
    PL4: '',
    VK11PT: '',
    VK11DB: '',
    VK11FOB: '',
    VK11TNPP: '',
    PTS: '',
    VK11BB: '',
    TH: '',
  }
  searchTermInput: { [key: string]: string } = {

    inputPrice: '',
    market: '',
    customerDb: '',
    customerPt: '',
    customerFob: '',
    customerTnpp: '',
    customerBbdo: '',
  }
  currentTab = ''
  currentTabInput = ''
  smsName = ''
  lstgoods: any[] = []
  lstSendSms: any[] = []
  localResult: any[] = []
  isBrowser: boolean = true;
  idramdom = new Date().getTime();
  isVisiblePreviewExcel: boolean = false
  urlViewExcel = ''

  constructor(
    private _service: CalculateDiscountService,
    private _localService : LocalService,
    private globalService: GlobalService,
    private message: NzMessageService,
    private route: ActivatedRoute,
    private _signerService: SignerService,
    private _configTemplateService: ConfigTemplateService,
    private _goodService: GoodsService,

    private location: Location
    // @Inject(PLATFORM_ID) private platformId: Object
  ) {
    // this.isBrowser = isPlatformBrowser(this.platformId);
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
        this.getOutput(this.headerId)
      },
    })
    this._service.getInput(this.headerId).subscribe({
      next: (data) => {
        this.input = data
        this.listNameBBDO = this.input.customerBbdo

        this.titleTab = data.header.name
      },
      error: (response) => {
        console.log(response)
      },
    })
    this.getRight()
    this.getAllLocal()
  }
  Showmessage(message: string, type: string): void {
    this.message.create(type, message)
  }

  getOutput(id: any) {
    this._service.getOutput(id).subscribe({
      next: (data) => {
        this.output = data

      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  getAllLocal() {
    this._localService.getall().subscribe({
      next: (data) => {
        this.localResult = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  getOutputCal(id: any) {
    this._service.getOutput(id).subscribe({
      next: (data) => {
        this.output = data
        this.Showmessage('Tính toán lại thành công', 'success')
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





  removeHtmlTags(html: string): string {
    if (!html) return ''
    let result = html.replace(/<[^>]*>/g, '');
    result = result.replace(/&nbsp;/g, ' ')  // Thay thế &nbsp; bằng khoảng trắng thường
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>');
    return result
  }

  // confirmSendsMail() {
  //   this._service.SendMail(this.headerId).subscribe({
  //     next: (data) => {
  //       this.message.create('success', 'Gửi mail thành công')
  //     },
  //     error: (err) => {
  //       console.log(err)
  //     },
  //   })
  // }
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
    })
  }
  getRight() {
    const rights = localStorage.getItem('userRights')
    this.rightList = rights ? JSON.parse(rights) : []

    const accountGroups = localStorage.getItem('UserInfo')
    this.accountGroups = accountGroups
      ? JSON.parse(accountGroups).accountGroups[0].name
      : []
    this.accountGroups == 'G_NV_K' ? (this.currentTab = 'PT') : ''
  }

  handleCancel() {
    this.isVisibleHistory = false
    this.isVisibleEmail = false
    this.isVisibleSms = false
    this.isVisibleExport = false
    this.isVisibleCustomer = false
    this.isVisibleCustomerPDF = false
    this.isSms = false
    this.checkedSms = false
    this.isVisiblePreviewExcel = false
    this.isVisiblePreview = false
    this.inputSearchCustomer = ""
  }
  exportWordTrinhKy() {
    this.isVisibleLstTrinhKy = !this.isVisibleLstTrinhKy
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
  }

  onAllCheckedLstTrinhKy(value: boolean): void {
    this.lstTrinhKyChecked = []
    if (value) {
      if (this.input.header.status != '04') {
        this.lstTrinhKy.forEach((i) => {
          if (i.status) {
            this.lstTrinhKyChecked.push(i.code)
          }
        })

      } else {
        this.lstTrinhKy.forEach((i) => {
          this.lstTrinhKyChecked.push(i.code)
        })
      }
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
            for (let index = 0; index < data.length; index++) {
              this.openNewTab(environment.apiUrl + data[index])
            }
            this.message.create('success', 'Xuất file thành công')
          },
          error: (err) => {
            console.log(err)
          },
        })
      this.lstTrinhKyChecked = []
      this.checked = false;
    }
  }

  onAllChecked(value: boolean): void {
    this.lstCustomerChecked = []
    if (value) {
      this.lstCustomer.forEach((i) => {
        this.lstCustomerChecked.push({
          code: i.code,
          deliveryGroupCode: i.deliveryGroupCode,
        })
      })
    } else {
      this.lstCustomerChecked = []
    }
  }







  checkedSms: boolean = false
  lstSearchSms: any[] = []

  showSMSAction() {
    this._service.GetSms(this.headerId).subscribe({
      next: (data) => {
        this.lstSearchSms = data
        this.lstSMS = data
        this.isVisibleSms = true
      },
      error: (err) => {
        console.log(err)
      },
    })
  }

  confirmSendSMS(smsname: any) {
    this._service.SaveSMS(this.headerId, smsname).subscribe({
      next: (data) => {
        console.log(data);

        if (data == "02") {
          this.message.create('warning', 'SMS thông báo giá bán lẻ đã đươc tạo!')
          return
        } else if (data == "01") {
          this.message.create('warning', 'SMS thông báo thù lao đã được tạo!')
          return
        }

        this.message.create('success', 'Tạo hàng chờ SMS thành công')
        this.showSMSAction()
      },
      error: (err) => {
        console.log(err)
      },
    })
  }

  ResendSms() {
    this._service.ResetSendSms(this.lstSendSmsChecked).subscribe({
      next: (data) => {
        this.isVisibleEmail = false

        this.message.create('success', 'Gửi lại mail thành công')
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  onAllCheckedSendSms(value: boolean): void {
    this.lstSendSmsChecked = []
    if (value) {
      this.lstSearchSms.forEach((i) => {
        if (i.isSend != "Y") {
          this.lstSendSmsChecked.push(i.id)
        }
      })
    } else {
      this.lstSendSmsChecked = []
    }
  }

  updateCheckedSetSendSms(code: any, checked: boolean,): void {
    if (checked) {
      this.lstSendSmsChecked.push(code)

    } else {
      this.lstSendSmsChecked = this.lstSendSmsChecked.filter(
        (x) => x !== code
      )
    }
  }

  onItemCheckedSendSms(code: String, checked: boolean,): void {
    this.updateCheckedSetSendSms(code, checked)
  }

  isCheckedSendSms(code: string): boolean {
    return this.lstSendSmsChecked.some((item) => item == code)
  }

  onSendSms() {
    if (this.lstSendSmsChecked.length == 0) {
      this.message.create(
        'warning',
        'Vui lòng chọn tin nhắn muốn gửi',
      )
      return
    } else {
      this._service.SendSMS(this.lstSendSmsChecked).subscribe({
        next: (data) => {
          this.lstSendSmsChecked = []
          this.handleCancel()
        }
      })
    }
  }

  selectedMarket: any = null
  selectedCustomer: any = null
  selectedTrangThai: any = null
  searchHistorySMS() {
    console.log(this.selectedCustomer, this.selectedMarket, this.inputSearchCustomer);

    this.lstSearchSms = this.lstSMS
    if (this.selectedMarket !== null || this.selectedCustomer !== null || this.inputSearchCustomer !== "" || this.selectedTrangThai !== null) {
      if (this.selectedMarket !== null) {
        this.lstSearchSms = this.lstSearchSms.filter(c =>
          c.marketCode == this.selectedMarket);
      }
      if (this.selectedCustomer !== null) {
        this.lstSearchSms = this.lstSearchSms.filter(c =>
          c.customerCode == this.selectedCustomer);
      }
      if (this.selectedTrangThai !== null) {
        this.lstSearchSms = this.lstSearchSms.filter(c =>{
          if (this.selectedTrangThai === "TB") {
              return c.isSend === "N" && c.numberRetry === 3;
            } else {
              return c.isSend === this.selectedTrangThai;
            }
          })
      }
      if (this.inputSearchCustomer !== "") {
        const keyword = this.inputSearchCustomer.trim().toLowerCase();
        this.lstSearchSms = this.lstSearchSms.filter(c =>
          c.contents.toLowerCase().includes(keyword) || c.phoneNumber.toLowerCase().includes(keyword))
      }
    } else {
      this.lstSearchSms = this.lstSMS
    }

  }

  clearSearchSms() {
    this.selectedMarket = null
    this.selectedCustomer = null
    this.selectedTrangThai = null
    this.inputSearchCustomer = ""
    this.searchHistorySMS()
  }




  checkedEmail: boolean = false
  lstSearchEmail: any[] = []
  showEmailAction() {
    this._service.Getmail(this.headerId).subscribe({
      next: (data) => {
        this.lstSearchEmail = data
        this.lstEmail = data
        this.isVisibleEmail = true
      },
      error: (err) => {
        console.log(err)
      },
    })
  }

  onAllCheckedSendEmail(value: boolean): void {
    this.lstSendEmailChecked = []
    if (value) {
      this.lstSearchEmail.forEach((i) => {
        if (i.isSend != "Y") {
          this.lstSendEmailChecked.push(i.id)
        }
      })
    } else {
      this.lstSendEmailChecked = []
    }
  }

  updateCheckedSetSendEmail(code: any, checked: boolean,): void {
    if (checked) {
      this.lstSendEmailChecked.push(code)

    } else {
      this.lstSendEmailChecked = this.lstSendEmailChecked.filter(
        (x) => x !== code
      )
    }
  }

  onItemCheckedSendEmail(code: String, checked: boolean,): void {
    this.updateCheckedSetSendEmail(code, checked)
  }

  isCheckedSendEmail(code: string): boolean {
    return this.lstSendEmailChecked.some((item) => item == code)
  }

  onSendEmail() {
    if (this.lstSendEmailChecked.length == 0) {
      this.message.create(
        'warning',
        'Vui lòng chọn tin nhắn muốn gửi',
      )
      return
    } else {
      this._service.SendlstEmail(this.lstSendEmailChecked).subscribe({
        next: () => {
          this.lstSendEmailChecked = []
          this.handleCancel()
        }
      })
    }
  }

  searchEmail() {
    const keyword = this.inputSearchMail.trim().toLowerCase();
    this.lstSearchEmail = this.lstEmail.filter(c =>
      c.contents.toLowerCase().includes(keyword)
    );
  }

  ResendEmail() {
    this._service.ResetSendlstMail(this.lstSendEmailChecked).subscribe({
      next: (data) => {
        this.isVisibleEmail = false

        this.message.create('success', 'Gửi lại mail thành công')
      },
      error: (response) => {
        console.log(response)
      },
    })
  }


  // selectedMarket: any = null
  selectedCustomerMail: any = null
  selectedTrangThaiMail : any = null
  searchHistoryMail() {
    this.lstSearchEmail = this.lstEmail
    if (this.selectedCustomerMail !== null || this.inputSearchCustomer !== "" || this.selectedTrangThaiMail !== null) {
      if (this.selectedCustomerMail !== null) {
        this.lstSearchEmail = this.lstSearchEmail.filter(c =>
          c.customerCode == this.selectedCustomerMail);
      }
      if (this.selectedTrangThaiMail !== null) {
        this.lstSearchEmail = this.lstSearchEmail.filter(c =>{
          if (this.selectedTrangThaiMail === "TB") {
              return c.isSend === "N" && c.numberRetry === 3;
            } else {
              return c.isSend === this.selectedTrangThaiMail;
            }
          }
        );
      }
      if (this.inputSearchCustomer !== "") {
        const keyword = this.inputSearchCustomer.trim().toLowerCase();
        this.lstSearchEmail = this.lstSearchEmail.filter(c =>
          c.contents.toLowerCase().includes(keyword) || c.email.toLowerCase().includes(keyword))
      }
    } else {
      this.lstSearchEmail = this.lstEmail
    }

  }

  clearSearchMail() {
    this.selectedCustomerMail = null
    this.selectedTrangThai = null
    this.inputSearchCustomer = ""
    this.searchHistoryMail()
  }


  searchTableBBDO() {
    const keyword = this.inputnameBBDO.toLowerCase();
    this.listNameBBDO = this.input.customerBbdo.filter((item: any) =>
      item.name.toLowerCase().includes(keyword)
    );
  }


  confirmExportWord() {
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
            this.message.create('success', 'Xuất file thành công')
          },
          error: (err) => {
            console.log(err)
          },
        })
    }
    this.lstCustomerChecked = [];
    this.checked = false;
  }
  confirmExportWordMail() {
    if (this.lstCustomerChecked.length == 0) {
      this.message.create(
        'warning',
        'Vui lòng chọn file cần tạo',
      )
      return
    } else {
      this._service
        .ExportWordMail(this.lstCustomerChecked, this.headerId)
        .subscribe({
          next: (data) => {
            this.isVisibleCustomer = false
            this.lstCustomerChecked = []

            this.message.create('success', 'Xuất file thành công')
          },
          error: (err) => {
            console.log(err)
          },
        })
    }
    this.lstCustomerChecked = [];
    this.checked = false;
  }

  updateCheckedSet(code: any, deliveryGroupCode: string, checked: boolean,): void {
    if (checked) {
      this.lstCustomerChecked.push({
        code: code,
        deliveryGroupCode: deliveryGroupCode,
      })
    } else {
      this.lstCustomerChecked = this.lstCustomerChecked.filter(
        (x) => x.code !== code || x.deliveryGroupCode !== deliveryGroupCode,
      )
    }
  }

  onItemChecked(code: String, deliveryGroupCode: string, checked: boolean,): void {
    this.updateCheckedSet(code, deliveryGroupCode, checked)
  }

  onClickTab(title: string, tab: number) {
    this.titleTab = title
  }


  // 01 khởi tạo || trình duyệt giá bán lẻ (01)
  // 02 chờ phê duyệt giá bán lẻ || hủy trình duyệt giá bán lẻ(02) , phê duyệt giá bán lẻ, yêu cầu chỉnh sửa
  // 03 yêu cầu chỉnh sửa giá bán lẻ || trình duyệt giá bán lẻ
  // 04 phê duyệt giá bán lẻ || hủy phê duyệt giá bán lẻ || trình duyệt giá thù lao
  // 05 từ chối ||
  // 06 chờ phê duyệt giá thù lao|| hủy trình duyệt giá thù lao, phê duyệt giá thù lao, yêu cầu chỉnh sửa
  // 07 yêu cầu chỉnh sửa giá thù lao || trình duyệt giá thù lao
  // 08 phê duyệt || phê duyệt giá thù lao

  changeStatus(value: string, status: string) {
    switch (value) {
      case '01':
        this.statusModel.title = 'TRÌNH DUYỆT'
        this.statusModel.des = 'Bạn có muốn Trình duyệt dữ liệu này?'
        // this.input.header.status = '01'/
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

      // case '06':
      //   this.statusModel.title = 'HỦY PHÊ DUYỆT GIÁ BÁN LẺ'
      //   this.statusModel.des = 'Bạn có muốn Hủy phê duyệt dữ liệu này?'
      //   break

      // case '11':
      //   this.statusModel.title = 'TRÌNH DUYỆT GIÁ THÙ LAO'
      //   this.statusModel.des = 'Bạn có muốn Trình duyệt dữ liệu này?'
      //   // this.input.header.status = '01'
      //   break
      // case '12':
      //   this.statusModel.title = 'YÊU CẦU CHỈNH SỬA GIÁ THÙ LAO'
      //   this.statusModel.des = 'Bạn có muốn Yêu cầu chỉnh sửa lại dữ liệu này?'
      //   break
      // case '13':
      //   this.statusModel.title = 'PHÊ DUYỆT'
      //   this.statusModel.des = 'Bạn có muốn Phê duyệt dữ liệu này?'
      //   break
      // case '15':
      //   this.statusModel.title = 'HỦY TRÌNH DUYỆT GIÁ THÙ LAO'
      //   this.statusModel.des = 'Bạn có muốn Hủy trình duyệt dữ liệu này?'
      //   break
      // case '16':
      //   this.statusModel.title = 'HỦY PHÊ DUYỆT'
      //   this.statusModel.des = 'Bạn có muốn Hủy phê duyệt dữ liệu này?'
      //   break
    }
    this.dataQuyTrinh.status.code = status
    this.dataQuyTrinh.header = this.input.header

    this.dataQuyTrinh.status.Link = window.location.href
    this.isVisibleStatus = true
    Swal.fire({
      title: this.statusModel.title,
      text: this.statusModel.des,
      input: 'text',
      inputPlaceholder: 'Ý kiến',
      icon: 'question',
      showCancelButton: true,
      confirmButtonText: 'Đồng ý',
      cancelButtonText: 'Hủy',
    }).then((result) => {
      if (result.isConfirmed) {
        this.dataQuyTrinh.status.content = result.value
        this._service.HandleQuyTrinh(this.dataQuyTrinh).subscribe({
          next: (data) => {
            window.location.reload()
          },
        })
      }
    })
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
    this.input2 = structuredClone(this.input)
    this.formatVcfAndBvmtData()
    this.visibleInput = true
    this.getAllSigner()
  }

  onUpdateInput() {
    this.visibleInput = false
    this._service.updateInput(this.input).subscribe({
      next: (data) => { },
      error: (response) => {
        console.log(response)
      },
    })
    this.getOutput(this.headerId)

  }

  handleQuyTrinh() {
    this._service.HandleQuyTrinh(this.input).subscribe({
      next: (data) => { },
      error: (response) => {
        console.log(response)
      },
    })
  }

  close(): void {
    this.visibleInput = false
  }

  reCalculate() {
    this.getOutputCal(this.headerId)
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
        this.isVisibleHistory = true
      },
      error: (err) => {
        console.log(err)
      },
    })
  }
  getCustomerName(code: string): string {
    if (code != null) {
      const customer = this.output.summary.find((c: any) => c.col4 === code);
      return customer ? customer.customerName : '';
    }
    return ""
  }

  getMarketName(code: string): string {
    const market = this.input.market.find((c: any) => c.code === code);
    return market ? market.name : '';
  }
  openNewTab(url: string) {
    window.open(url, '_blank')
  }

  exportExcel() {
    this._service.exportExcel(this.headerId, this.accountGroups).subscribe({
      next: (data) => {
        var a = document.createElement('a')
        a.href = environment.apiUrl + data
        a.target = '_blank'
        a.click()
        a.remove()
        this.message.create('success', 'Xuất file thành công')
      },
      error: (response) => {
        console.log(response)
      },
    })
  }


  Preview(data: any) {

    if (data.type == "xlsx") {
      this.urlViewExcel = `http://sso.d2s.com.vn:1235/${data.path}?cacheBuster=${new Date().getTime()}`
      this.isVisiblePreviewExcel = true
      this.config = {
        document: {
          fileType: 'xlsx',
          key: `ket${this.idramdom}`,
          title: 'Example Document Title.xlsx',
          url: `${this.urlViewExcel}`,
        },
        documentType: 'cell',
        editorConfig: {
          mode: 'view',

        },
      };
    } else if (data.type == "docx") {
      this.urlViewExcel = `http://sso.d2s.com.vn:1235/${data.path}?cacheBuster=${new Date().getTime()}`
      this.isVisiblePreviewExcel = true
      this.config = {
        document: {
          fileType: 'docx',
          key: `keydoc${this.idramdom}`,
          title: 'File.docx',
          url: `${this.urlViewExcel}`,
        },
        documentType: 'word',
        editorConfig: {
          mode: 'view',

        },
      };

    }

  }

  config: IConfig = {
    document: {
      fileType: 'xlsx',
      key: `ket${this.idramdom}`,
      title: 'Example Document Title.xlsx',
      url: `${this.urlViewExcel}`,
    },
    documentType: 'cell',
    editorConfig: {
      mode: 'view',

    },
  };


  onShowSMS() {
    this._configTemplateService.getall().subscribe({
      next: (data) => {
        this.lstSms = data.filter((item: any) => item.type === "SMS")
      },
      error: (response) => {
        console.log(response)
      },
    })
    this.isSms = true
  }

  isAllCheckedSmsFirstChange: boolean = false
  allChecked = false;


  onInputNumberFormat(data: any, field: string) {
    let value = data[field]
    // 1. Bỏ ký tự không hợp lệ (chỉ giữ số, '-', '.')
    value = value.replace(/[^0-9\-.]/g, '')

    // 2. Đảm bảo chỉ có 1 dấu '-' và nó đứng đầu
    const minusMatches = value.match(/-/g)
    if (minusMatches && minusMatches.length > 1) {
      value = value.replace(/-/g, '') // Xoá hết
      value = '-' + value // Thêm 1 dấu '-' đầu tiên
    } else if (minusMatches && !value.startsWith('-')) {
      value = value.replace(/-/g, '')
      value = '-' + value
    }

    // 3. Xử lý dấu '.': chỉ cho sau '0' hoặc '-0' và duy nhất
    const dotIndex = value.indexOf('.')
    if (dotIndex !== -1) {
      const beforeDot = value.substring(0, dotIndex)
      const afterDot = value.substring(dotIndex + 1).replace(/\./g, '')

      if (beforeDot === '0' || beforeDot === '-0') {
        value = beforeDot + '.' + afterDot
      } else {
        // Loại bỏ dấu '.' nếu không đúng điều kiện
        value = beforeDot + afterDot
      }
    }

    // 4. Format phần nguyên với dấu ','
    const parts = value.split('.')
    let integerPart = parts[0].replace(/[^0-9\-]/g, '') // giữ dấu '-'
    integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')

    // 5. Ghép lại
    let formattedValue = integerPart
    if (parts[1]) {
      formattedValue += '.' + parts[1]
    }

    // 6. Cập nhật lại giá trị hiển thị
    data[field] = formattedValue
    // 7. Parse về số
    const rawNumber = formattedValue.replace(/,/g, '')
    const numberValue = parseFloat(rawNumber)
    const finalNumber = isNaN(numberValue) ? 0 : numberValue
    // 8. Update vào model chuẩn
    const index = this.input2.inputPrice.findIndex(
      (x: any) => x.goodCode === data.goodCode,
    )
    if (index !== -1) {
      this.input.inputPrice[index][field] = finalNumber
    }
  }

  onKeyDownNumberOnly(event: KeyboardEvent) {
    const allowedKeys = [
      'Backspace',
      'ArrowLeft',
      'ArrowRight',
      'Delete',
      'Tab',
      '-',
      '.',
    ]

    // Cho phép dùng Ctrl/Cmd kết hợp với: A, C, V, X
    if (
      (event.ctrlKey || event.metaKey) &&
      ['a', 'c', 'v', 'x', 'z'].includes(event.key.toLowerCase())
    ) {
      return
    }

    // Cho phép nhập số hoặc phím được phép
    if (
      (event.key >= '0' && event.key <= '9') ||
      allowedKeys.includes(event.key)
    ) {
      return
    }

    // Chặn các phím còn lại
    event.preventDefault()
  }

  handleAutoInput(row: any) {
    const index = this.input2.inputPrice.indexOf(row)

    this.input2.inputPrice[index].fobV1 = parseInt(this.input2.inputPrice[index].fobV2.replace(/,/g, ''), 10)
    this.input.inputPrice[index].fobV1 = this.input2.inputPrice[index].fobV1
    this.input2.inputPrice[index].fobV1 = this.formatNumber(this.input2.inputPrice[index].fobV1)

  }

  formatNumber(value: any): string {
    if (value == null || value === '') return ''

    const num = parseFloat(value.toString().replace(/,/g, ''))
    if (isNaN(num)) return ''

    // Format giữ 4 chữ số sau dấu phẩy (mày có thể chỉnh lại tuỳ)
    return num.toLocaleString('en-US', {
      minimumFractionDigits: 0,
      maximumFractionDigits: 4,
    })
  }

  formatVcfAndBvmtData() {
    if (this.input2.inputPrice && Array.isArray(this.input2.inputPrice)) {
      this.input2.inputPrice.forEach((item: any) => {
        // Format các trường số cần format
        item.vcf = this.formatNumber(item.vcf)
        item.thueBvmt = this.formatNumber(item.thueBvmt)
        item.chenhLech = this.formatNumber(item.chenhLech)
        item.gblV1 = this.formatNumber(item.gblV1)
        item.gblV2 = this.formatNumber(item.gblV2)
        item.l15Blv2 = this.formatNumber(item.l15Blv2)
        item.l15Nbl = this.formatNumber(item.l15Nbl)
        item.laiGop = this.formatNumber(item.laiGop)
        item.fobV1 = this.formatNumber(item.fobV1)
        item.fobV2 = this.formatNumber(item.fobV2)
      })
    }
  }


  onDocumentReady = () => {
    if (this.isBrowser) {
      console.log('Document is loaded');
    }
  };

  onLoadComponentError = (errorCode: number, errorDescription: string) => {
    if (this.isBrowser) {
      switch (errorCode) {
        case -1: // Unknown error loading component
          console.log(errorDescription);
          break;
        case -2: // Error loading DocsAPI
          console.log(errorDescription);
          break;
        case -3: // DocsAPI is not defined
          console.log(errorDescription);
          break;
      }
    }
  };

  exportWord() {
    this._service.GetCustomerBbdo(this.headerId).subscribe({
      next: (data) => {
        this.lstCustomer = data
        this.lstCus = data
        this.isVisibleCustomer = true

      },
    })
  }
  isChecked(code: string): boolean {
    return this.lstCustomerChecked.some((item) => item.code === code)
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
    this.lstCustomerChecked = []
    this.checked = false;
  }
  exportPDF() {
    this._service.GetCustomerBbdo(this.headerId).subscribe({
      next: (data) => {
        this.lstCustomer = data
        this.isVisibleCustomerPDF = true
      },
    })
  }
  search(sheetName: string) {
    this.searchTerm[sheetName] = this.searchInput
  }
  searchInPutDb(sheetName: string) {
    if (sheetName == "") {
      sheetName = 'inputPrice'
    }
    this.searchTermInput[sheetName] = this.searchInputTab

  }

  reset(tabName: string) {
    const keys = Object.keys(this.searchTerm)
    keys.forEach((key) => (this.searchTerm[key] = ''))
    this.searchInput = ''
    this.currentTab = tabName
  }
  resetInput(tabName: string) {
    const keys = Object.keys(this.searchTermInput)
    keys.forEach((key) => (this.searchTermInput[key] = ''))
    this.searchInputTab = ''
    this.currentTabInput = tabName

  }

  getSearchTerm(key: string): string {
    return this.searchTerm[key] || ''
  }
  getSearchTermInput(key: string): string {
    return this.searchTermInput[key] || ''
  }

  hasExportPermission(): boolean {
    return (
      this.rightList.includes(IMPORT_BATCH.EXPORT_TO_EXCEL) ||
      this.rightList.includes(IMPORT_BATCH.EXPORT_TO_PDF) ||
      this.rightList.includes(IMPORT_BATCH.EXPORT_TO_WORD)
    )
  }
  onDateChange(date: Date) {
    const month = new Date(date).getMonth() + 1
    // Tạo array mới để Angular detect thay đổi
    const temp = this.input.inputPrice.map((item1: any) => {
      const matched = this.lstgoods.find(
        (item2) => item2.code === item1.goodCode,
      )
      const vcfValue = matched
        ? month >= 5 && month <= 10
          ? matched.vfcHt
          : matched.vfcDx
        : item1.vcf

      return { ...item1, vcf: vcfValue } // tạo object mới luôn
    })

    this.input.inputPrice = temp
    this.input2.inputPrice = structuredClone(temp)

    this.formatVcfAndBvmtData()
  }

  formatNegativeNumber(value: number | null | undefined): string {
    if (value == null || isNaN(value)) return '';

    const formatted = Math.abs(value).toLocaleString('en-US'); // format dấu phẩy phần nghìn

    return value < 0
      ? `(${formatted})`
      : formatted;
  }

  formatNegativeNumber2(value: number | null | undefined): string {
    if (value == null || isNaN(value)) return '';

    const roundedValue = Math.round(value); // Làm tròn tới đơn vị
    const formatted = Math.abs(roundedValue).toLocaleString('en-US'); // Định dạng dấu phẩy phần nghìn

    return roundedValue < 0 ? `(${formatted})` : formatted;
  }

  lstCus: any[] = []
  searchCustomer() {
    const keyword = this.inputSearchCustomer.trim().toLowerCase();
    this.lstCus = this.lstCustomer.filter(c =>
      c.name.toLowerCase().includes(keyword)
    );
  }



}
