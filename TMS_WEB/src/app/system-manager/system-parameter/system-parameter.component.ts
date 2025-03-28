import {Component} from '@angular/core';
import {ShareModule} from '../../shared/share-module';
import {GlobalService} from '../../services/global.service';
import {PaginationResult} from '../../models/base.model';
import {FormControl, FormGroup, Validators, NonNullableFormBuilder} from '@angular/forms';
import {NzFormLayoutType} from 'ng-zorro-antd/form';
import {SystemParameterFilter} from '../../models/system-manager/system-parameter.model';
import {SystemParamaterService} from '../../services/system-manager/system-parameter.service';

@Component({
  selector: 'app-system-parameter',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './system-parameter.component.html',
  styleUrl: './system-parameter.component.scss',
})
export class SystemParameterComponent {
  validateForm: FormGroup = this.fb.group({
    code: ['', [Validators.required]],
    name: ['', [Validators.required]],
    value: ['', [Validators.required]],
  });
  loading: boolean = false;
  isSubmit: boolean = false;
  visible: boolean = false;
  edit: boolean = false;
  filter = new SystemParameterFilter();

  paginationResult = new PaginationResult();

  constructor(
    private _service: SystemParamaterService,
    private fb: NonNullableFormBuilder,
    private globalService: GlobalService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách cấu hình tham số hệ thống',
        path: 'system-manager/system-parameter',
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
    this.search();
  }

  onSortChange(name: string, value: any) {
    this.filter = {
      ...this.filter,
      SortColumn: name,
      IsDescending: value === 'descend',
    };
    this.search();
  }

  search() {
    this.isSubmit = false;
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
    return this._service.exportExcel(this.filter).subscribe((result: Blob) => {
      const blob = new Blob([result], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const url = window.URL.createObjectURL(blob);
      var anchor = document.createElement('a');
      anchor.download = 'danh-sach-cau-hinh-he-thong.xlsx';
      anchor.href = url;
      anchor.click();
    });
  }

  submitForm(): void {
    this.isSubmit = true;
    if (this.validateForm.valid) {
      if (this.edit) {
        this._service.update(this.validateForm.getRawValue()).subscribe({
          next: (data) => {
            this.search();
          },
          error: (response) => {
            console.log(response);
          },
        });
      } else {
        this._service.create(this.validateForm.getRawValue()).subscribe({
          next: (data) => {
            this.search();
          },
          error: (response) => {
            console.log(response);
          },
        });
      }
    } else {
      Object.values(this.validateForm.controls).forEach((control) => {
        if (control.invalid) {
          control.markAsDirty();
          control.updateValueAndValidity({onlySelf: true});
        }
      });
    }
  }

  close() {
    this.visible = false;
    this.resetForm();
  }

  reset() {
    this.filter = new SystemParameterFilter();
    this.search();
  }

  openCreate() {
    this.edit = false;
    this.visible = true;
  }

  resetForm() {
    this.validateForm.reset();
  }

  deleteItem(code: string | number) {
    this._service.delete(code).subscribe({
      next: (data) => {
        this.search();
      },
      error: (response) => {
        console.log(response);
      },
    });
  }

  openEdit(data: {code: string; name: string; value: string}) {
    this.edit = true;
    this.visible = true;
    this.validateForm.setValue({
      code: data?.code,
      name: data.name,
      value: data.value,
    });
  }

  pageSizeChange(size: number): void {
    this.filter.currentPage = 1;
    this.filter.pageSize = size;
    this.search();
  }

  pageIndexChange(index: number): void {
    this.filter.currentPage = index;
    this.search();
  }
}
