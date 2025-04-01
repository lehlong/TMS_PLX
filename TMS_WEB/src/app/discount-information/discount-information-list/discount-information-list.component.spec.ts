import { ComponentFixture, TestBed } from '@angular/core/testing';

import { DiscountInformationListComponent } from './discount-information-list.component';

describe('DiscountInformationListComponent', () => {
  let component: DiscountInformationListComponent;
  let fixture: ComponentFixture<DiscountInformationListComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DiscountInformationListComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(DiscountInformationListComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
