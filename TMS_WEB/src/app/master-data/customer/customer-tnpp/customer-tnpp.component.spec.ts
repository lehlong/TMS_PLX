import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CustomerTnppComponent } from './customer-tnpp.component';

describe('CustomerTnppComponent', () => {
  let component: CustomerTnppComponent;
  let fixture: ComponentFixture<CustomerTnppComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CustomerTnppComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CustomerTnppComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
