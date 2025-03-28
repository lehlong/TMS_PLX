import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CalculateDiscountComponent } from './calculate-discount.component';

describe('CalculateDiscountComponent', () => {
  let component: CalculateDiscountComponent;
  let fixture: ComponentFixture<CalculateDiscountComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CalculateDiscountComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CalculateDiscountComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
