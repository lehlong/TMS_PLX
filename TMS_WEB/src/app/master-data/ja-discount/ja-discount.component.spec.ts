import { ComponentFixture, TestBed } from '@angular/core/testing';

import { JaDiscountComponent } from './ja-discount.component';

describe('JaDiscountComponent', () => {
  let component: JaDiscountComponent;
  let fixture: ComponentFixture<JaDiscountComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [JaDiscountComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(JaDiscountComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
