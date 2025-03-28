import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CustomerPtComponent } from './customer-pt.component';

describe('CustomerPtComponent', () => {
  let component: CustomerPtComponent;
  let fixture: ComponentFixture<CustomerPtComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CustomerPtComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CustomerPtComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
