import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CustomerBbdoComponent } from './customer-bbdo.component';

describe('CustomerBbdoComponent', () => {
  let component: CustomerBbdoComponent;
  let fixture: ComponentFixture<CustomerBbdoComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CustomerBbdoComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CustomerBbdoComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
