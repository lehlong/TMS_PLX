import { ComponentFixture, TestBed } from '@angular/core/testing';

import { TermOfPaymentComponent } from './term-of-payment.component';

describe('TermOfPaymentComponent', () => {
  let component: TermOfPaymentComponent;
  let fixture: ComponentFixture<TermOfPaymentComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [TermOfPaymentComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(TermOfPaymentComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
