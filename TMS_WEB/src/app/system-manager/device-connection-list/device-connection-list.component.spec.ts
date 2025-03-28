import { ComponentFixture, TestBed } from '@angular/core/testing';

import { DeviceConnectionListComponent } from './device-connection-list.component';

describe('DeviceConnectionListComponent', () => {
  let component: DeviceConnectionListComponent;
  let fixture: ComponentFixture<DeviceConnectionListComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DeviceConnectionListComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(DeviceConnectionListComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
