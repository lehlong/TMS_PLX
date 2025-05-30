import { Component, OnInit } from '@angular/core'
import { ActivatedRoute } from '@angular/router'
import { CalculateDiscountService } from '../../services/calculate-discount/calculate-discount.service'
import { GlobalService } from '../../services/global.service'
import { ShareModule } from '../../shared/share-module'
import { environment } from '../../../environments/environment.prod'
import { NgxDocViewerModule } from 'ngx-doc-viewer'
import { IMPORT_BATCH } from '../../shared/constants/access-right.constants'
import { SignerService } from '../../services/master-data/signer.service'
import Swal from 'sweetalert2'
import { GoodsService } from '../../services/master-data/goods.service'
import { DocumentEditorModule } from '@onlyoffice/document-editor-angular'
import { IConfig } from '@onlyoffice/document-editor-angular'
import { NzMessageService } from 'ng-zorro-antd/message'
import { ConfigTemplateService } from '../../services/system-manager/config-template.service'
import { LocalService } from '../../services/master-data/local.service'
import { lstTrinhKy } from '../../shared/constants/select.model'
import { OutputModel } from '../../models/calculate-discount/output.model'
import { InputModel } from '../../models/calculate-discount/input.model'

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

  dataQuyTrinh: any = {
    header: {},
    status: {},
  }
  input2: any = new InputModel()
  rightList: any = []
  statusModel: any = {
    title: '',
    des: '',
    value: '',
  }

  selectedIndexInput: any = 0
  onTabChangeInput(e: any) {
    this.selectedIndexInput = e
  }

  getFilteredList(): any[] {
    const term = this.getSearchTermInput('listNameBBDO');
    if (!term) return this.listNameBBDO;
    return this.listNameBBDO.filter(item =>
      item?.name?.toLowerCase().includes(term)
    );
  }

  trackByCode(index: number, item: any): string {
    return item.code;
  }

  input: any = new InputModel()
  output: any = new OutputModel()
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
  searchValue = ''
  visible = false
  listOfData: any[] = []
  lstTrinhKy = lstTrinhKy
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
  isConfirmLoading = false
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
  isBrowser: boolean = true
  idramdom = new Date().getTime()
  isVisiblePreviewExcel: boolean = false
  urlViewExcel = ''

  constructor(
    private _service: CalculateDiscountService,
    private _localService: LocalService,
    public _global: GlobalService,
    private message: NzMessageService,
    private route: ActivatedRoute,
    private _signerService: SignerService,
    private _configTemplateService: ConfigTemplateService,
    private _goodService: GoodsService,
  ) {
    this._global.setBreadcrumb([
      {
        name: 'Kết quả tính toán',
        path: 'calculate-discount/detail',
      },
    ])
    this._global.getLoading().subscribe((value) => {
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

    this.getMasterData()
  }

  getMasterData() {
    const rights = localStorage.getItem('userRights')
    this.rightList = rights ? JSON.parse(rights) : []

    const accountGroups = localStorage.getItem('UserInfo')
    this.accountGroups = accountGroups
      ? JSON.parse(accountGroups).accountGroups[0].name
      : []
    this.accountGroups == 'G_NV_K' ? (this.currentTab = 'PT') : ''

    this._localService.getall().subscribe((data) => (this.localResult = data))
    this._goodService.getall().subscribe((data) => (this.lstgoods = data))
    this._signerService.getall().subscribe((data) => (this.signerResult = data))
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
    this.inputSearchCustomer = ''
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
      this.checked = false
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
        console.log(data)

        if (data == '02') {
          this.message.create(
            'warning',
            'SMS thông báo giá bán lẻ đã đươc tạo!',
          )
          return
        } else if (data == '01') {
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
        if (i.isSend != 'Y') {
          this.lstSendSmsChecked.push(i.id)
        }
      })
    } else {
      this.lstSendSmsChecked = []
    }
  }

  updateCheckedSetSendSms(code: any, checked: boolean): void {
    if (checked) {
      this.lstSendSmsChecked.push(code)
    } else {
      this.lstSendSmsChecked = this.lstSendSmsChecked.filter((x) => x !== code)
    }
  }

  onItemCheckedSendSms(code: String, checked: boolean): void {
    this.updateCheckedSetSendSms(code, checked)
  }

  isCheckedSendSms(code: string): boolean {
    return this.lstSendSmsChecked.some((item) => item == code)
  }

  onSendSms() {
    if (this.lstSendSmsChecked.length == 0) {
      this.message.create('warning', 'Vui lòng chọn tin nhắn muốn gửi')
      return
    } else {
      this._service.SendSMS(this.lstSendSmsChecked).subscribe({
        next: (data) => {
          this.lstSendSmsChecked = []
          this.handleCancel()
        },
      })
    }
  }

  selectedMarket: any = null
  selectedCustomer: any = null
  selectedTrangThai: any = null

  searchHistorySMS() {
    const keyword = this.inputSearchCustomer.trim().toLowerCase()
    this.lstSearchSms = this.lstSMS
      .filter((c) => !this.selectedMarket || c.marketCode === this.selectedMarket,)
      .filter((c) => !this.selectedCustomer || c.customerCode === this.selectedCustomer,)
      .filter((c) => {
        if (!this.selectedTrangThai) return true
        return this.selectedTrangThai === 'TB'
          ? c.isSend === 'N' && c.numberRetry === 3/*  */
          : c.isSend === this.selectedTrangThai
      })
      .filter((c) =>
        !keyword ||
        c.contents.toLowerCase().includes(keyword) ||
        c.phoneNumber.toLowerCase().includes(keyword),
      )
  }

  clearSearchSms() {
    this.selectedMarket = null
    this.selectedCustomer = null
    this.selectedTrangThai = null
    this.inputSearchCustomer = ''
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
        if (i.isSend != 'Y') {
          this.lstSendEmailChecked.push(i.id)
        }
      })
    } else {
      this.lstSendEmailChecked = []
    }
  }

  updateCheckedSetSendEmail(code: any, checked: boolean): void {
    if (checked) {
      this.lstSendEmailChecked.push(code)
    } else {
      this.lstSendEmailChecked = this.lstSendEmailChecked.filter(
        (x) => x !== code,
      )
    }
  }

  onItemCheckedSendEmail(code: String, checked: boolean): void {
    this.updateCheckedSetSendEmail(code, checked)
  }

  isCheckedSendEmail(code: string): boolean {
    return this.lstSendEmailChecked.some((item) => item == code)
  }

  onSendEmail() {
    if (this.lstSendEmailChecked.length == 0) {
      this.message.create('warning', 'Vui lòng chọn tin nhắn muốn gửi')
      return
    } else {
      this._service.SendlstEmail(this.lstSendEmailChecked).subscribe({
        next: () => {
          this.lstSendEmailChecked = []
          this.handleCancel()
        },
      })
    }
  }

  searchEmail() {
    const keyword = this.inputSearchMail.trim().toLowerCase()
    this.lstSearchEmail = this.lstEmail.filter((c) =>
      c.contents.toLowerCase().includes(keyword),
    )
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

  selectedCustomerMail: any = null
  selectedTrangThaiMail: any = null
  searchHistoryMail() {
    const keyword = this.inputSearchCustomer.trim().toLowerCase()

    this.lstSearchEmail = this.lstEmail
      .filter((c) =>
        !this.selectedCustomerMail ||
        c.customerCode === this.selectedCustomerMail,
      )
      .filter((c) => {
        if (!this.selectedTrangThaiMail) return true
        return this.selectedTrangThaiMail === 'TB'
          ? c.isSend === 'N' && c.numberRetry === 3
          : c.isSend === this.selectedTrangThaiMail
      })
      .filter(
        (c) =>
          !keyword ||
          c.contents.toLowerCase().includes(keyword) ||
          c.email.toLowerCase().includes(keyword),
      )
  }

  clearSearchMail() {
    this.selectedCustomerMail = null
    this.selectedTrangThai = null
    this.inputSearchCustomer = ''
    this.searchHistoryMail()
  }


  lstCus: any[] = []
  searchDelivery: any = ''
  searchCustomer() {
    const keyword = this.inputSearchCustomer.trim().toLowerCase()
    this.lstCus = this.lstCustomer
    .filter((c) => !keyword || c.name.toLowerCase().includes(keyword))
    .filter((c) => !this.searchDelivery || c.deliveryGroupCode == this.searchDelivery)
    console.log(keyword, this.lstCus)
  }


  searchTableBBDO() {
    const keyword = this.inputnameBBDO.toLowerCase()
    this.listNameBBDO = this.input.customerBbdo.filter((item: any) =>
      item.name.toLowerCase().includes(keyword),
    )
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
    this.lstCustomerChecked = []
    this.checked = false
  }
  confirmExportWordMail() {
    if (this.lstCustomerChecked.length == 0) {
      this.message.create('warning', 'Vui lòng chọn file cần tạo')
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
    this.lstCustomerChecked = []
    this.checked = false
  }

  updateCheckedSet(
    code: any,
    deliveryGroupCode: string,
    checked: boolean,
  ): void {
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

  onItemChecked(
    code: String,
    deliveryGroupCode: string,
    checked: boolean,
  ): void {
    this.updateCheckedSet(code, deliveryGroupCode, checked)
  }

  onClickTab(title: string, tab: number) {
    this.titleTab = title
  }

  changeStatus(value: string, status: string) {
    const statusMap: Record<string, { title: string; des: string }> = {
      '01': {
        title: 'TRÌNH DUYỆT',
        des: 'Bạn có muốn Trình duyệt dữ liệu này?',
      },
      '02': {
        title: 'YÊU CẦU CHỈNH SỬA',
        des: 'Bạn có muốn Yêu cầu chỉnh sửa lại dữ liệu này?',
      },
      '03': { title: 'PHÊ DUYỆT', des: 'Bạn có muốn Phê duyệt dữ liệu này?' },
      '04': { title: 'TỪ CHỐI', des: 'Bạn có muốn Từ chối dữ liệu này?' },
      '05': {
        title: 'HỦY TRÌNH DUYỆT',
        des: 'Bạn có muốn Hủy trình duyệt dữ liệu này?',
      },
      '06': {
        title: 'HỦY PHÊ DUYỆT',
        des: 'Bạn có muốn Hủy phê duyệt dữ liệu này?',
      },
    }

    const selectedStatus = statusMap[value]
    if (!selectedStatus) return

    this.statusModel = { ...selectedStatus }

    this.dataQuyTrinh.status = {
      code: status,
      Link: window.location.href,
    }
    this.dataQuyTrinh.header = this.input.header

    this.isVisibleStatus = true

    Swal.fire({
      title: selectedStatus.title,
      text: selectedStatus.des,
      input: 'text',
      inputPlaceholder: 'Ý kiến',
      icon: 'question',
      showCancelButton: true,
      confirmButtonText: 'Đồng ý',
      cancelButtonText: 'Hủy',
    }).then((result) => {
      if (result.isConfirmed) {
        this.dataQuyTrinh.status.content = result.value
        this._service.HandleQuyTrinh(this.dataQuyTrinh).subscribe(() => {
          window.location.reload()
        })
      }
    })
  }

  openInput() {
    this.input2 = structuredClone(this.input)
    this.formatVcfAndBvmtData()
    this.visibleInput = true
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
    this.ngOnInit()
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
      const customer = this.output.summary.find((c: any) => c.col4 === code)
      return customer ? customer.customerName : ''
    }
    return ''
  }

  getMarketName(code: string): string {
    const market = this.input.market.find((c: any) => c.code === code)
    return market ? market.name : ''
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
    if (!['xlsx', 'docx'].includes(data.type)) return
    const fileType = data.type
    const title =
      fileType === 'xlsx' ? 'Example Document Title.xlsx' : 'File.docx'
    const key = (fileType === 'xlsx' ? 'ket' : 'keydoc') + this.idramdom
    const documentType = fileType === 'xlsx' ? 'cell' : 'word'
    this.urlViewExcel = `http://sso.d2s.com.vn:1235/${data.path}?cacheBuster=${Date.now()}`
    this.isVisiblePreviewExcel = true
    this.config = {
      document: {
        fileType,
        key,
        title,
        url: this.urlViewExcel,
      },
      documentType,
      editorConfig: {
        mode: 'view',
      },
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
  }

  onShowSMS() {
    this._configTemplateService.getall().subscribe({
      next: (data) => {
        this.lstSms = data.filter((item: any) => item.type === 'SMS')
      },
      error: (response) => {
        console.log(response)
      },
    })
    this.isSms = true
  }

  isAllCheckedSmsFirstChange: boolean = false
  allChecked = false

  onInputNumberFormat(data: any, field: string) {
    let value = data[field]?.toString().replace(/[^0-9.-]/g, '') || ''

    value = value.replace(/-/g, '')
    if (data[field].startsWith('-')) value = '-' + value

    const [intPartRaw, ...decParts] = value.split('.')
    const intPart = intPartRaw.replace(/[^0-9-]/g, '')
    const decPart = decParts.join('').replace(/\./g, '')

    const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
    const formattedValue = decPart ? `${formattedInt}.${decPart}` : formattedInt

    data[field] = formattedValue

    const numberValue = parseFloat(formattedValue.replace(/,/g, '')) || 0

    const index = this.input2.inputPrice.findIndex(
      (x: any) => x.goodCode === data.goodCode,
    )
    if (index !== -1) {
      this.input.inputPrice[index][field] = numberValue
    }
  }

  onKeyDownNumberOnly(event: KeyboardEvent) {
    const { key, ctrlKey, metaKey } = event
    const isModifierKey = ctrlKey || metaKey
    const allowedKeys = [
      'Backspace',
      'ArrowLeft',
      'ArrowRight',
      'Delete',
      'Tab',
      '-',
      '.',
    ]
    const isCtrlCombo =
      isModifierKey && ['a', 'c', 'v', 'x', 'z'].includes(key.toLowerCase())
    const isDigit = key >= '0' && key <= '9'
    if (isDigit || allowedKeys.includes(key) || isCtrlCombo) return
    event.preventDefault()
  }

  handleAutoInput(row: any) {
    const index = this.input2.inputPrice.indexOf(row)
    this.input2.inputPrice[index].fobV1 = parseInt(
      this.input2.inputPrice[index].fobV2.replace(/,/g, ''),
      10,
    )
    this.input.inputPrice[index].fobV1 = this.input2.inputPrice[index].fobV1
    this.input2.inputPrice[index].fobV1 = this._global.formatNumber(
      this.input2.inputPrice[index].fobV1,
    )
  }

  formatVcfAndBvmtData() {
    if (this.input2.inputPrice && Array.isArray(this.input2.inputPrice)) {
      this.input2.inputPrice.forEach((item: any) => {
        item.vcf = this._global.formatNumber(item.vcf)
        item.thueBvmt = this._global.formatNumber(item.thueBvmt)
        item.chenhLech = this._global.formatNumber(item.chenhLech)
        item.gblV1 = this._global.formatNumber(item.gblV1)
        item.gblV2 = this._global.formatNumber(item.gblV2)
        item.l15Blv2 = this._global.formatNumber(item.l15Blv2)
        item.l15Nbl = this._global.formatNumber(item.l15Nbl)
        item.laiGop = this._global.formatNumber(item.laiGop)
        item.fobV1 = this._global.formatNumber(item.fobV1)
        item.fobV2 = this._global.formatNumber(item.fobV2)
      })
    }
  }

  onDocumentReady = () => {
    if (this.isBrowser) {
      console.log('Document is loaded')
    }
  }

  onLoadComponentError = (errorCode: number, errorDescription: string) => {
    if (this.isBrowser) {
      switch (errorCode) {
        case -1:
          console.log(errorDescription)
          break
        case -2:
          console.log(errorDescription)
          break
        case -3:
          console.log(errorDescription)
          break
      }
    }
  }

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
    if (!this.lstCustomerChecked.length) {
      this.message.create(
        'warning',
        'Vui lòng chọn khách hàng cần xuất ra file',
      )
      return
    }
    this._service.ExportPDF(this.lstCustomerChecked, this.headerId).subscribe({
      next: (data) => {
        this.isVisibleCustomer = false
        this.lstCustomerChecked = []
        this.checked = false
        const link = document.createElement('a')
        link.href = `${environment.apiUrl}${data}`
        link.target = '_blank'
        link.click()
        link.remove()
      },
      error: (err) => console.error(err),
    })
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
    if (sheetName == '') {
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
    const isHtSeason = date.getMonth() + 1 >= 5 && date.getMonth() + 1 <= 10
    const updatedPrices = this.input.inputPrice.map((item: any) => {
      const matched = this.lstgoods.find((g) => g.code === item.goodCode)
      return {
        ...item,
        vcf: matched ? (isHtSeason ? matched.vfcHt : matched.vfcDx) : item.vcf,
      }
    })
    this.input.inputPrice = updatedPrices
    this.input2.inputPrice = structuredClone(updatedPrices)
    this.formatVcfAndBvmtData()
  }

}
