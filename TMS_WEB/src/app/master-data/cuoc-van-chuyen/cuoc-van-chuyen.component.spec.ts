import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CuocVanChuyenComponent } from './cuoc-van-chuyen.component';

describe('CuocVanChuyenComponent', () => {
  let component: CuocVanChuyenComponent;
  let fixture: ComponentFixture<CuocVanChuyenComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CuocVanChuyenComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CuocVanChuyenComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
