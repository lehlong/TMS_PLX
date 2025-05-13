import { ComponentFixture, TestBed } from '@angular/core/testing';

import { JaPtPhanPhoiComponent } from './ja-pt-phan-phoi.component';

describe('JaPtPhanPhoiComponent', () => {
  let component: JaPtPhanPhoiComponent;
  let fixture: ComponentFixture<JaPtPhanPhoiComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [JaPtPhanPhoiComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(JaPtPhanPhoiComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
