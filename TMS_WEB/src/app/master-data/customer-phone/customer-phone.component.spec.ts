import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CustomerPhoneComponent } from './customer-phone.component';

describe('CustomerPhoneComponent', () => {
  let component: CustomerPhoneComponent;
  let fixture: ComponentFixture<CustomerPhoneComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CustomerPhoneComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CustomerPhoneComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
