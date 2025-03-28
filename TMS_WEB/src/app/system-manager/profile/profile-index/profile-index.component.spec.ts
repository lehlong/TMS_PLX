import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ProfileIndexComponent } from './profile-index.component';

describe('ProfileIndexComponent', () => {
  let component: ProfileIndexComponent;
  let fixture: ComponentFixture<ProfileIndexComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ProfileIndexComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(ProfileIndexComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
