import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ConfigMailComponent } from './config-mail.component';

describe('ConfigMailComponent', () => {
  let component: ConfigMailComponent;
  let fixture: ComponentFixture<ConfigMailComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ConfigMailComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(ConfigMailComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
