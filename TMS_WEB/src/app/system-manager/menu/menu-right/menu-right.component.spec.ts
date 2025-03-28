import { ComponentFixture, TestBed } from '@angular/core/testing';

import { MenuRightComponent } from './menu-right.component';

describe('MenuRightComponent', () => {
  let component: MenuRightComponent;
  let fixture: ComponentFixture<MenuRightComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [MenuRightComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(MenuRightComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
