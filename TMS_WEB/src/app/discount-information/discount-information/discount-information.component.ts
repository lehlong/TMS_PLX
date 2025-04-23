import { Component } from '@angular/core';
import { GlobalService } from '../../services/global.service';
import { ShareModule } from '../../shared/share-module';
import { ActivatedRoute } from '@angular/router';
import { GoodsService } from '../../services/master-data/goods.service';
import { DiscountInformationService } from '../../services/discount-information/discount-information.service';
import { DiscountInformationListService } from '../../services/discount-information/discount-information-list.service';
import { environment } from '../../../environments/environment.prod';
import { COMPETITOR_ANALYSIS } from '../../shared/constants';
import { CalculateDiscountService } from '../../services/calculate-discount/calculate-discount.service';

@Component({
  selector: 'app-discount-information',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './discount-information.component.html',
  styleUrl: './discount-information.component.scss'
})
export class DiscountInformationComponent {
  COMPETITOR_ANALYSIS = COMPETITOR_ANALYSIS
  constructor(
    private _service: DiscountInformationService,
    private _serviceCD: CalculateDiscountService,
    private _discountInformationList: DiscountInformationListService,
    private globalService: GlobalService,
    private route: ActivatedRoute,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Kết quả phân tích',
        path: 'calculate-result',
      },
    ])
    this.globalService.getLoading().subscribe((value) => {
      this.loading = value
    })
  }


  loading: boolean = false
  edit: boolean = true
  visible: boolean = false
  isName: boolean = false
  isVisibleExport: boolean = false
  headerName: any = 'THÔNG TIN PHÂN TÍCH ĐỢT NHẬP - '

  code: any = ''
  title: any = 'Phân tích chiết khấu'
  name: any = ''
  headerId: any = ''
  fDate: any = ''
  lstHistoryFile : any[] = []
  data: any = {
    lstDIL: [{}],
    lstGoods : [],
    lstCompetitor : [],
    discount: []
  }
  model: any = {
    goodss: [{
      code: '',
      hs: [],
      discountCompany:[],
      discount:[]
    }],
    header: {
      name: '',
      fData: ''
    },
    headerName: ''
  }

  ngOnInit() {
    this.route.paramMap.subscribe({
      next: (params) => {
        this.code = params.get('code')
        this.getAll()
      },
    })
  }

  getAll() {
    this._service.getAll(this.code).subscribe({
      next: (data) => {
        this.data = data
        this.name = this.data.lstDIL.name ?? ""
        this.headerId = this.data.lstDIL.code
        this.fDate = new Date(this.data.lstDIL.fDate).toLocaleDateString()
        this.changeTitle(this.fDate)
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  submitForm(): void {
    console.log(this.model)
    var m = {
      model: this.model
    }
    this._discountInformationList.createData(this.model).subscribe({
      next: (data) => {
        console.log(data)
      }
    })
  }

  updateDataInput() {
    if (this.model.header.name != ''){
      this._service.UpdateDataInput(this.model).subscribe({
        next: (data) => {
          window.location.reload()
        },
        error: (err) => {
          console.log(err)
        },
      })
    }
  }

  changeTitle(fDate: string){
    this.title = 'Phân tích chiết khấu ngày ' + fDate
    this.headerName = this.headerName + fDate
  }

  getDataHeader(){
    this._service.getDataInput(this.code).subscribe({
      next: (data) => {
        this.visible = true
        this.model = data
      },
    })
  }

  reCalculate(){
    this.getAll()
  }

  showHistoryExport(){
    this._serviceCD.GetHistoryFile(this.headerId).subscribe({
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

  openNewTab(url: string) {
    console.log(url);
    window.open(url, '_blank')
  }


  exportExcel(){
    this._service.ExportExcel(this.code).subscribe({
      next: (data) => {
        var a = document.createElement('a')
        a.href = environment.apiUrl + data
        a.target = '_blank'
        a.click()
        a.remove()
      },
    })
  }

  exportExcelBaoCaoThuLao(){
    this._service.ExportExcelBaoCaoThuLao(this.code).subscribe({
      next: (data) => {
        var a = document.createElement('a')
        a.href = environment.apiUrl + data
        a.target = '_blank'
        a.click()
        a.remove()
      },
    })
  }

  checkName(_name: string){
    _name == '' ? this.isName = true : this.isName = false
  }

  close() {
    this.headerName = 'THÔNG TIN PHÂN TÍCH ĐỢT NHẬP - '
    this.visible = false
    this.isVisibleExport = false
  }




}
