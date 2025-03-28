import {Injectable} from '@angular/core';
import {AbstractControl, ValidationErrors} from '@angular/forms';
import {Router, Route} from '@angular/router';
@Injectable({
  providedIn: 'root',
})
export class utils {
  constructor(private router: Router) {}
  trimSpace(control: AbstractControl): ValidationErrors | null {
    if (control.value && control.value.toString().trim() == '') {
      return {cannotContainSpace: true};
    }
    return null;
  }

  formatNumber(value: string | number): string {
    if (value == null) return '';

    if (typeof value === 'number') {
      if (value % 1 === 0) {
        return value.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      } else {
        return value.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      }
    } else if (typeof value !== 'string') {
      return '';
    }
    const [integerPart, decimalPart] = value.split('.');
    const formattedIntegerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    if (decimalPart === undefined) {
      return formattedIntegerPart;
    }
    return formattedIntegerPart + '.' + decimalPart;
  }
}