import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CustomerDbComponent } from './customer-db.component';

describe('CustomerDbComponent', () => {
  let component: CustomerDbComponent;
  let fixture: ComponentFixture<CustomerDbComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CustomerDbComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CustomerDbComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
