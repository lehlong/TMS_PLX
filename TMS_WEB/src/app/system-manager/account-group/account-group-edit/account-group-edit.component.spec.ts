import { ComponentFixture, TestBed } from '@angular/core/testing';

import { AccountGroupEditComponent } from './account-group-edit.component';

describe('AccountGroupEditComponent', () => {
  let component: AccountGroupEditComponent;
  let fixture: ComponentFixture<AccountGroupEditComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [AccountGroupEditComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(AccountGroupEditComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
