import { Component, OnInit } from '@angular/core';
import { ShareModule } from '../../shared/share-module';
import { CalculateDiscountService } from '../../services/calculate-discount/calculate-discount.service';
import { GlobalService } from '../../services/global.service';
import { NzMessageService } from 'ng-zorro-antd/message';
import { BaseFilter, PaginationResult } from '../../models/base.model';
import { Router } from '@angular/router';
import { FormControl } from '@angular/forms'
import { SignerService } from '../../services/master-data/signer.service';
import { GoodsService } from '../../services/master-data/goods.service';

@Component({
  selector: 'app-calculate-discount',
  standalone: true,
  imports: [ShareModule],
  templateUrl: './calculate-discount.component.html',
  styleUrl: './calculate-discount.component.scss'
})
export class CalculateDiscountComponent implements OnInit {
  constructor(
    private _service: CalculateDiscountService,
    private globalService: GlobalService,
    private message: NzMessageService,
    private router: Router,
    private _signerService: SignerService,
    private _goodService: GoodsService,
  ) {
    this.globalService.setBreadcrumb([
      {
        name: 'Danh sách đợt tính mức giảm giá',
        path: 'calculate-discount/list',
      },
    ])
    this.globalService.getLoading().subscribe((value) => {
      this.loading = value
    })
  }
  isVisibleStatus: boolean = false
  noData: any[] = []
  loading: boolean = false
  visible = false;
  filter = new BaseFilter()
  paginationResult = new PaginationResult()


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
  input2: any = this.input;
  nguoiKyControl = new FormControl({ code: "", name: "", position: "" });
  signerResult: any[] = []
  selectedValue = {}
  lstgoods: any[] = []
  ngOnInit(): void {
    this.search();
    this.getAllSigner();
    this.getAllGood();
  }
  search() {
    this._service.search(this.filter).subscribe({
      next: (data) => {
        this.paginationResult = data
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

  genarateCreate() {
    this._service.genarateCreate().subscribe({
      next: (data) => {
        this.input = data;
        const month = new Date(this.input.header.date).getMonth() + 1;
        this.input.inputPrice.forEach((item1: any) => {
          const matched = this.lstgoods.find(item2 => item2.code === item1.goodCode);
          if (matched) {
            item1.vcf = (month >= 5 && month <= 10) ? matched.vfcHt : matched.vfcDx;
          }
        });

        this.input2 = structuredClone(data)
        this.formatVcfAndBvmtData()
        this.visible = true;
      },
      error: (response) => {
        console.log(response)
      },
    })
  }
  getInputUpdate(headerId: any) {
    this._service.copyInput(headerId, this.input.header.id).subscribe({
      next: (data) => {
        this.input = data
        this.input2 = structuredClone(data)
        this.formatVcfAndBvmtData()
        // this.titleTab = data.header.name
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  getAllSigner() {
    this._signerService.getall().subscribe({
      next: (data) => {
        this.signerResult = data
        this.selectedValue = this.signerResult.find(item => item.code === "d72636e2-454f-4085-b491-76b2e0c6445d");
      },
      error: (response) => {
        console.log(response)
      },
    })
  }

  onCreate() {
    this.input.header.signerCode = this.nguoiKyControl.value?.code || ''
    // console.log(this.input);

    this._service.create(this.input).subscribe({
      next: (data) => {
        this.router.navigate([`/calculate-discount/detail/${this.input.header.id}`]);
      },
      error: (response) => {
        console.log(response)
      },
    });
  }
  openCalculateDetail(id: any) {
    this.router.navigate([`/calculate-discount/detail/${id}`]);
  }
  close(): void {
    this.visible = false;
    this.input = {
      header: {},
      inputPrice: [],
      market: [],
      customerDb: [],
      customerPt: [],
      customerFob: [],
      customerTnpp: [],
    };
  }
  onChange(result: Date): void {
    //console.log('Selected Time: ', result);
  }

  onOk(result: Date | Date[] | null): void {
    //console.log('onOk', result);
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
      this.input2.inputPrice.forEach((item: any) => {
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
  onDateChange(date: Date) {
    const month = new Date(date).getMonth() + 1;

    // Tạo array mới để Angular detect thay đổi
    const temp = this.input.inputPrice.map((item1: any) => {
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


  handleAutoInput(row: any) {
    const index = this.input2.inputPrice.indexOf(row)

    this.input2.inputPrice[index].fobV1 = parseInt(this.input2.inputPrice[index].fobV2.replace(/,/g, ''), 10)
    this.input.inputPrice[index].fobV1 = this.input2.inputPrice[index].fobV1
    this.input2.inputPrice[index].fobV1 = this.formatNumber(this.input2.inputPrice[index].fobV1)
    console.log(this.input2.inputPrice[index].fobV1);

  }


}
