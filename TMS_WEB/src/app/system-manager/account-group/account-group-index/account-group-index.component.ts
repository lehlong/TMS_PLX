import {Component, ViewChild} from '@angular/core';
import {ShareModule} from '../../../shared/share-module';
import {AccountGroupFilter} from '../../../models/system-manager/account-group.model';
import {GlobalService} from '../../../services/global.service';
import {AccountGroupService} from '../../../services/system-manager/account-group.service';
import {PaginationResult} from '../../../models/base.model';
import {AccountGroupCreateComponent} from '../account-group-create/account-group-create.component';
import {AccountGroupEditComponent} from '../account-group-edit/account-group-edit.component';
import { ACCOUNT_GROUP_RIGHTS, ADMIN_RIGHTS } from '../../../shared/constants';
@Component({
  selector: 'app-account-group-index',
  standalone: true,
  imports: [ShareModule, AccountGroupCreateComponent, AccountGroupEditComponent],
  templateUrl: './account-group-index.component.html',
  styleUrl: './account-group-index.component.scss',
})
export class AccountGroupIndexComponent {
  @ViewChild(AccountGroupEditComponent) accountGroupEditComponent!: AccountGroupEditComponent;

  filter = new AccountGroupFilter();
  paginationResult = new PaginationResult();
  showCreate: boolean = false;
  showEdit: boolean = false;
  idDetail: number | string = 0;
  loading: boolean = false;
  ADMIN_RIGHTS = ADMIN_RIGHTS
  constructor(private _service: AccountGroupService, private globalService: GlobalService) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách nhóm tài khoản',
        path: 'system-manager/account-group',
      },
    ]);
    this.globalService.getLoading().subscribe((value) => {
      this.loading = value;
    });
  }

  ngOnDestroy() {
    this.globalService.setBreadcrumb([]);
  }

  ngOnInit(): void {
    this.loadInit();
  }

  loadInit() {
    this.search();
  }

  openCreate() {
    this.showCreate = true;
  }

  onSortChange(name: string, value: any) {
    this.filter = {
      ...this.filter,
      SortColumn: name,
      IsDescending: value === 'descend',
    };
    this.search();
  }
  openEdit(id: number | string) {
    this.idDetail = id;
    this.showEdit = true;
    this.accountGroupEditComponent.loadDetail(this.idDetail);
  }

  close() {
    this.showCreate = false;
    this.showEdit = false;
  }

  search() {
    this._service.search(this.filter).subscribe({
      next: (data) => {
        this.paginationResult = data;
      },
      error: (response) => {
        console.log(response);
      },
    });
  }

  exportExcel() {
    return this._service.ExportExcel(this.filter).subscribe((result: Blob) => {
      const blob = new Blob([result], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const url = window.URL.createObjectURL(blob);
      var anchor = document.createElement('a');
      anchor.download = 'danh-sach-doi-tac.xlsx';
      anchor.href = url;
      anchor.click();
    });
  }

  reset() {
    this.filter = new AccountGroupFilter();
    this.loadInit();
  }

  pageSizeChange(size: number): void {
    this.filter.currentPage = 1;
    this.filter.pageSize = size;
  }

  pageIndexChange(index: number): void {
    this.filter.currentPage = index;
  }

  deleteItem(id: string | number) {
    this._service.delete(id).subscribe({
      next: (data) => {
        this.search();
      },
      error: (response) => {
        console.log(response);
      },
    });
  }
}
