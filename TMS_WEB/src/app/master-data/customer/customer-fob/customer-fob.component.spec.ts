import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CustomerFobComponent } from './customer-fob.component';

describe('CustomerFobComponent', () => {
  let component: CustomerFobComponent;
  let fixture: ComponentFixture<CustomerFobComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CustomerFobComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CustomerFobComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
