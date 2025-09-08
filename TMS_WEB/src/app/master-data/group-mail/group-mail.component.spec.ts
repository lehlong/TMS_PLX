import { ComponentFixture, TestBed } from '@angular/core/testing';

import { GroupMailComponent } from './group-mail.component';

describe('GroupMailComponent', () => {
  let component: GroupMailComponent;
  let fixture: ComponentFixture<GroupMailComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [GroupMailComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(GroupMailComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
