import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CalculateDiscountDetailComponent } from './calculate-discount-detail.component';

describe('CalculateDiscountDetailComponent', () => {
  let component: CalculateDiscountDetailComponent;
  let fixture: ComponentFixture<CalculateDiscountDetailComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CalculateDiscountDetailComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CalculateDiscountDetailComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
