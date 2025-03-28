import { ComponentFixture, TestBed } from '@angular/core/testing';

import { AccountGroupCreateComponent } from './account-group-create.component';

describe('AccountGroupCreateComponent', () => {
  let component: AccountGroupCreateComponent;
  let fixture: ComponentFixture<AccountGroupCreateComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [AccountGroupCreateComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(AccountGroupCreateComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
