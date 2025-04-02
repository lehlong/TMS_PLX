import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CustomerPtsComponent } from './customer-pts.component';

describe('CustomerPtsComponent', () => {
  let component: CustomerPtsComponent;
  let fixture: ComponentFixture<CustomerPtsComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CustomerPtsComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CustomerPtsComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
