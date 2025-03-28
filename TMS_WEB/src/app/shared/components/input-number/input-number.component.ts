import {Component, Input, forwardRef, Output, EventEmitter} from '@angular/core';
import {ControlValueAccessor, NG_VALUE_ACCESSOR, FormControl} from '@angular/forms';
import {utils} from '../../utils/index';
import {InputClearComponent} from '../input-clear/input-clear.component';
import {ReactiveFormsModule} from '@angular/forms';
import {CommonModule} from '@angular/common';
import {FormsModule} from '@angular/forms';
import {NzFormModule} from 'ng-zorro-antd/form';
import {NzAlertModule} from 'ng-zorro-antd/alert';

@Component({
  selector: 'app-input-number',
  templateUrl: './input-number.component.html',
  styleUrls: ['./input-number.component.scss'],
  standalone: true,
  imports: [ReactiveFormsModule, CommonModule, InputClearComponent, FormsModule, NzFormModule, NzAlertModule],
  providers: [
    {
      provide: NG_VALUE_ACCESSOR,
      useExisting: forwardRef(() => InputNumberComponent),
      multi: true,
    },
  ],
})
export class InputNumberComponent implements ControlValueAccessor {
  @Input() focus: boolean = false;
  @Input() notHandle: boolean = false;
  @Input() initFocus: boolean = false;
  @Output() input = new EventEmitter<string>();
  @Output() click = new EventEmitter<string>();
  @Input() control!: FormControl;
  @Input() class: string = '';
  @Input() classWraper: string = '';
  @Input() nzErrorTip: string = '';
  @Input() label: string = '';
  @Input() labelBold: boolean = false;
  @Input() required: boolean = false;
  @Input() placeholder: string = '';
  @Input() showErrors: any = null;
  @Input() errorsRequired: any = null;
  @Input() disabled: boolean = false;
  value: string = '';
  @Input() valueInput!: number | string;
  @Output() valueInputChange = new EventEmitter<number | string>();
  @Output() blurChange = new EventEmitter<number>();

  constructor(private utils: utils) {}

  onChange!: (value: any) => void;
  onTouched!: () => void;

  onBlur(): void {
    this?.blurChange?.emit();
  }

  onClick(e: any) {
    if (this.click) {
      this.click.emit();
    }
  }

  checkError(): boolean {
    if (this.control.errors) {
      const {required, email, pattern, ...errorOther} = this.control.errors;
      if (!required && !email && !pattern && errorOther && Object.keys(errorOther).length !== 0) {
        return true;
      }
    }
    return false;
  }

  writeValue(): void {
    this.value = this.utils.formatNumber(this.value);
    if (this.control) {
      this.control.setValue(
        !isNaN(parseFloat(this.value.replace(/,/g, ''))) ? parseFloat(this.value.replace(/,/g, '')) : '',
      );
    } else {
      this.valueInput = !isNaN(parseFloat(this.value.replace(/,/g, '')))
        ? parseFloat(this.value.replace(/,/g, ''))
        : '';
      this.valueInputChange.emit(this.valueInput);
    }
    if (this.input) {
      this.input.emit(this.value.replace(/,/g, ''));
    }
  }

  clearZero() {
    if (this.value == '0') {
      this.value = '';
    }
  }

  registerOnChange(fn: any): void {
    this.onChange = fn;
  }

  registerOnTouched(fn: any): void {
    this.onTouched = fn;
  }

  ngOnChanges(changes: any): void {
    if (this.notHandle) {
      this.value = this.utils.formatNumber(changes.valueInput.currentValue);
    } else {
      if (changes.valueInput && !changes.valueInput.firstChange) {
        this.value = this.utils.formatNumber(changes.valueInput.currentValue);
        this.valueInputChange.emit(changes.valueInput.currentValue);
      }
    }
  }

  ngOnInit() {
    if (this.control) {
      this.value = this.utils.formatNumber(this.control.value);
      this.control.valueChanges.subscribe((value: any) => {
        this.value = this.utils.formatNumber(value);
      });
    } else {
      this.value = this.utils.formatNumber(this.valueInput);
    }
  }
}
