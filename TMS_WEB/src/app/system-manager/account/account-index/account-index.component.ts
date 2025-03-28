import { Component, ViewChild } from '@angular/core'
import { ShareModule } from '../../../shared/share-module'
import { GlobalService } from '../../../services/global.service'
import { DropdownService } from '../../../services/dropdown/dropdown.service'
//import {PartnerManagementService} from '../../../services/business/partner-management.service';
import { PaginationResult } from '../../../models/base.model'
import { AccountService } from '../../../services/system-manager/account.service'
import { AccountCreateComponent } from '../account-create/account-create.component'
import { AccountFilter } from '../../../models/system-manager/account.model'
import { AccountEditComponent } from '../account-edit/account-edit.component'
import { AccountGroupEditComponent } from '../../account-group/account-group-edit/account-group-edit.component'
import { ActivatedRoute, Router } from '@angular/router'
import { ADMIN_RIGHTS } from '../../../shared/constants'

@Component({
  selector: 'app-account-index',
  standalone: true,
  imports: [
    ShareModule,
    AccountCreateComponent,
    AccountEditComponent,
    AccountGroupEditComponent,
  ],
  templateUrl: './account-index.component.html',
  styleUrl: './account-index.component.scss',
})
export class AccountIndexComponent {
  filter = new AccountFilter()
  paginationResult = new PaginationResult()
  showCreate: boolean = false
  showEdit: boolean = false
  userName: string = ''

  listAccountGroup: any[] = []
  //listPartner: any[] = []
  accountType: any[] = []
  listStatus: any[] = [
    { id: 'true', name: 'Kích hoạt' },
    { id: 'false', name: 'Khoá' },
  ]

  showEditAcg: boolean = false
  idDetail: number | string = 0
  // UserTypeCodes = UserTypeCodes

  @ViewChild(AccountEditComponent) accountEditComponent!: AccountEditComponent
  @ViewChild(AccountGroupEditComponent)
  accountGroupEditComponent!: AccountGroupEditComponent
  ADMIN_RIGHTS = ADMIN_RIGHTS
  constructor(
    private dropdownService: DropdownService,
    // private _service: PartnerManagementService,
    private _as: AccountService,
    private globalService: GlobalService,
    private route: ActivatedRoute,
    private router: Router,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách tài khoản',
        path: 'system-manager/account',
      },
    ])
  }

  ngOnDestroy() {
    this.globalService.setBreadcrumb([])
  }

  ngOnInit(): void {
    this.loadInit()
  }

  ngAfterViewInit() {
    this.route.queryParams.subscribe((params) => {
      if (params['user_name']) {
        this.openEdit(params['user_name'])
      }
      if (params['create_nmtv']) {
        this.openCreate()
      }
    })
  }

  onSortChange(key: string, value: string | null) {
    this.filter = {
      ...this.filter,
      SortColumn: key,
      IsDescending: value === 'descend',
    }
    this.search()
  }

  loadInit() {
    this.getAllAccountGroup()
    this.getAllAccountType()
    this.search()
  }

  openCreate() {
    this.showCreate = true
  }

  openEdit(userName: string) {
    this.userName = userName
    this.showEdit = true
    this.accountEditComponent?.getDetail(this.userName)
    const params = { user_name: userName }
    this.router.navigate(['/system-manager/account'], { queryParams: params })
    console.log('detail', this.accountEditComponent?.getDetail(this.userName))
  }

  close() {
    this.showCreate = false
    this.showEdit = false
    this.showEditAcg = false
    this.loadInit()
    // this.router.navigate([], {
    //   queryParams: {},
    // });
  }

  // getUserTypeText(type: string) {
  //   if (type === 'KH') return 'Khách hàng'
  //   else if (type === 'NM') return 'Nhà máy'
  //   else if (type === 'LX') return 'Lái xe'
  //   else if (type === 'NM_TV') return 'Nhân viên thương vụ'
  //   return ''
  // }

  search() {
    this._as
      .search({
        ...this.filter,
      })
      .subscribe({
        next: (data) => {
          this.paginationResult = data
        },
        error: (response) => {
          console.log(response)
        },
      })
  }
  getAllAccountGroup() {
    this.dropdownService.getAllAccountGroup().subscribe({
      next: (data) => {
        this.listAccountGroup = data
      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  getAllAccountType() {
    this.dropdownService.getAllAccountType().subscribe({
      next: (data) => {
        this.accountType = data
        console.log("AccountType", this.accountType);

      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  getAccountTypeNameById(id: string | number): string {
    const accountType = this.accountType.find(item => item.id === id);
    return accountType ? accountType.name : 'N/A';
  }

  // exportExcel() {
  //   return this._service.exportExcel(this.filter).subscribe((result: Blob) => {
  //     const blob = new Blob([result], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  //     const url = window.URL.createObjectURL(blob);
  //     var anchor = document.createElement('a');
  //     anchor.download = 'danh-sach-doi-tac.xlsx';
  //     anchor.href = url;
  //     anchor.click();
  //   });
  // }

  reset() {
    this.filter = new AccountFilter()
    this.loadInit()
  }

  pageSizeChange(size: number): void {
    this.filter.currentPage = 1
    this.filter.pageSize = size
    this.loadInit()
  }

  pageIndexChange(index: number): void {
    this.filter.currentPage = index
    this.loadInit()
  }

  handleAccountGroup(id: number | string) {
    this.idDetail = id
    this.showEditAcg = true
    this.accountGroupEditComponent.loadDetail(this.idDetail)
  }
}
