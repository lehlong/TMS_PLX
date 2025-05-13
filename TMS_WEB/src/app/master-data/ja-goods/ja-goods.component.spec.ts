import { ComponentFixture, TestBed } from '@angular/core/testing';

import { JaGoodsComponent } from './ja-goods.component';

describe('JaGoodsComponent', () => {
  let component: JaGoodsComponent;
  let fixture: ComponentFixture<JaGoodsComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [JaGoodsComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(JaGoodsComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
