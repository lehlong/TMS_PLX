import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ConfigSmsComponent } from './config-sms.component';

describe('ConfigSmsComponent', () => {
  let component: ConfigSmsComponent;
  let fixture: ComponentFixture<ConfigSmsComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ConfigSmsComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(ConfigSmsComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
