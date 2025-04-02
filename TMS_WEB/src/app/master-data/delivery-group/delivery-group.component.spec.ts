import { ComponentFixture, TestBed } from '@angular/core/testing';

import { DeliveryGroupComponent } from './delivery-group.component';

describe('DeliveryGroupComponent', () => {
  let component: DeliveryGroupComponent;
  let fixture: ComponentFixture<DeliveryGroupComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DeliveryGroupComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(DeliveryGroupComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
