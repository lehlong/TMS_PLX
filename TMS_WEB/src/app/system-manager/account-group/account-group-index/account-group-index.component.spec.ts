import { ComponentFixture, TestBed } from '@angular/core/testing';

import { AccountGroupIndexComponent } from './account-group-index.component';

describe('AccountGroupIndexComponent', () => {
  let component: AccountGroupIndexComponent;
  let fixture: ComponentFixture<AccountGroupIndexComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [AccountGroupIndexComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(AccountGroupIndexComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
