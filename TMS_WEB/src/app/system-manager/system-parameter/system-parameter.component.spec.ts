import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SystemParameterComponent } from './system-parameter.component';

describe('SystemParameterComponent', () => {
  let component: SystemParameterComponent;
  let fixture: ComponentFixture<SystemParameterComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [SystemParameterComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(SystemParameterComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
