import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CuocVanChuyenListComponent } from './cuoc-van-chuyen-list.component';

describe('CuocVanChuyenListComponent', () => {
  let component: CuocVanChuyenListComponent;
  let fixture: ComponentFixture<CuocVanChuyenListComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CuocVanChuyenListComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(CuocVanChuyenListComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
