import { ComponentFixture, TestBed } from '@angular/core/testing'

import { ConfixTemplateEmailComponent } from './config-template-email.component'

describe('ConfixTemplateEmailComponent', () => {
  let component: ConfixTemplateEmailComponent
  let fixture: ComponentFixture<ConfixTemplateEmailComponent>

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ConfixTemplateEmailComponent],
    }).compileComponents()

    fixture = TestBed.createComponent(ConfixTemplateEmailComponent)
    component = fixture.componentInstance
    fixture.detectChanges()
  })

  it('should create', () => {
    expect(component).toBeTruthy()
  })
})
