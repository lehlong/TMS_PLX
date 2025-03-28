import { Component, AfterViewInit, OnInit, OnDestroy, ChangeDetectorRef, ChangeDetectionStrategy } from '@angular/core';
import { ShareModule } from '../../shared/share-module';
import { GlobalService } from '../../services/global.service';
import { DeviceConnectionService } from '../../services/system-manager/device-connection.service';
import { PaginationResult } from '../../models/base.model';
import { DeviceConnectionFilter } from '../../models/system-manager/device-connection.model';
import { NzCardModule } from 'ng-zorro-antd/card';
import { FormGroup, Validators, NonNullableFormBuilder } from '@angular/forms';
import { format } from 'date-fns';

@Component({
  selector: 'app-device-connection-list',
  standalone: true,
  imports: [ShareModule, NzCardModule],
  templateUrl: './device-connection-list.component.html',
  styleUrls: ['./device-connection-list.component.scss'],
  changeDetection: ChangeDetectionStrategy.Default
})
export class DeviceConnectionListComponent implements AfterViewInit, OnInit, OnDestroy {
  validateForm: FormGroup = this.fb.group({
    code: ['', [Validators.required]],
    name: ['', [Validators.required]],
    type: ['', [Validators.required]],
    address: ['', [Validators.required]],
    interval: ['', [Validators.required]],
    note: [''],
  });

  paginationResult = new PaginationResult();
  showCreate = false;
  showEdit = false;
  idDetail: number | string = 0;
  loading = false;
  filter = new DeviceConnectionFilter();
  listDevice: any[] = [];
  edit = false;
  visible = false;
  isSubmit = false;
  isVisible = false; // Hiển thị modal
  modalContent = ''; // Nội dung của modal
  currentDeviceCode: string | number = '';
  textErrorOrderNumber = '';
  messageContent: any;

  constructor(
    private _service: DeviceConnectionService,
    private globalService: GlobalService,
    private fb: NonNullableFormBuilder,
    private cdr: ChangeDetectorRef,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách thiết bị',
        path: 'system-manager/device-connection',
      },
    ]);
    this.globalService.getLoading().subscribe((value) => {
      this.loading = value;
    });
    this._service.initiateSignalrConnection();
  }

  ngAfterViewInit() {
    this._service.hubDeviceConnection.subscribe((message: any) => {
      if (message && message !== '') {
        this.updateListDevice(message); // Cập nhật danh sách thiết bị
        this.updateModalContent(message); // Cập nhật modal
      }
    });
  }

  ngOnInit(): void {
    this.loadInit();
  }

  ngOnDestroy(): void {
    this.globalService.setBreadcrumb([]);
  }

  loadInit(): void {
    this.getAllDevice();
  }

  close() {
    this.visible = false;
    this.resetForm();
  }

  reset() {
    this.getAllDevice();
  }

  openCreate() {
    this.edit = false;
    this.visible = true;
  }

  resetForm() {
    this.validateForm.reset();
    this.textErrorOrderNumber = '';
    this.isSubmit = false;
  }

  openEdit(item: { code: string; name: string; address: string; type: string; interval: number; note: string }, event: MouseEvent) {
    event.stopPropagation();
    this.edit = true;
    this.visible = true;
    this.validateForm.setValue({
      code: item.code,
      name: item.name,
      type: item.type,
      address: item.address,
      interval: item.interval,
      note: item.note || '',
    });
  }

  getAllDevice(): void {
    this._service.getAllDevice().subscribe({
      next: (data) => {
        this.listDevice = [...data];
      },
      error: (response) => {
        console.log(response);
      },
    });
  }

  search(): void {}

  submitForm(): void {
    this.isSubmit = true;
    if (this.validateForm.valid) {
      if (this.edit) {
        this._service.Update(this.validateForm.getRawValue()).subscribe({
          next: () => {
            this.getAllDevice();
          },
          error: (response) => {
            console.log(response);
          },
        });
      } else {
        this._service.Insert(this.validateForm.getRawValue()).subscribe({
          next: () => {

              this.getAllDevice();

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
          control.updateValueAndValidity({ onlySelf: true });
        }
      });
    }
  }

  confirmDelete(code: string | number, event: MouseEvent) {
    event.stopPropagation();
  }

  deleteItem(code: string | number) {
    this._service.delete(code).subscribe({
      next: (data) => {
        this.getAllDevice();
      },
      error: (response) => {
        console.log(response);
      },
    });
  }

  openModal(item: any): void {
    this.currentDeviceCode = item.code;
    this.modalContent = this.getModalContent(item);
    this.isVisible = true;
  }

  handleCancel(): void {
    this.isVisible = false;
  }

  getModalContent(item: any): string {
    let statusColor = item.status ? 'green' : (item.status === false ? 'red' : 'gray');
    let statusText = item.status ? 'True' : (item.status === false ? 'False' : 'Unknown');

    // Format lastCheckTime
    let formattedLastCheckTime = format(new Date(item.lastCheckTime), 'dd/MM/yyyy HH:mm:ss');

    return `<p>Mã thiết bị: ${item.code}</p>
            <p>Status: <span style="color: ${statusColor}">${statusText}</span></p>
            <p>Log: ${item.log}</p>
            <p>Check Time: ${formattedLastCheckTime}</p>`;
  }

  updateModalContent(message: any): void {
    if (this.isVisible && this.currentDeviceCode === message.code) {
      const deviceIndex = this.listDevice.findIndex(device => device.code === message.code);
      console.log(deviceIndex);
      if (deviceIndex !== -1) {
        let item = this.listDevice[deviceIndex];
        if (message.checkTime) {
          item.lastCheckTime = message.checkTime;
        }
        if (message.message) {
          item.log = message.message;
        }
        this.modalContent = this.getModalContent(item);
        this.cdr.detectChanges();
      }
    }
  }

  updateListDevice(message: any) {
    const deviceIndex = this.listDevice.findIndex(device => device.code === message.code);
    if (deviceIndex !== -1) {
      Object.assign(this.listDevice[deviceIndex], message);
      this.cdr.detectChanges();
    } else {
      console.log('Device not found:', message.code);
    }
  }
}
