import { ComponentFixture, TestBed } from '@angular/core/testing';

import { JanaPriceComponent } from './jana-price.component';

describe('JanaPriceComponent', () => {
  let component: JanaPriceComponent;
  let fixture: ComponentFixture<JanaPriceComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [JanaPriceComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(JanaPriceComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
