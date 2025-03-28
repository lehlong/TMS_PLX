import { ComponentFixture, TestBed } from '@angular/core/testing'

import { ConfixTemplateSmsComponent } from './config-template-sms.component'

describe('ConfixTemplateSmsComponent', () => {
  let component: ConfixTemplateSmsComponent
  let fixture: ComponentFixture<ConfixTemplateSmsComponent>

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ConfixTemplateSmsComponent],
    }).compileComponents()

    fixture = TestBed.createComponent(ConfixTemplateSmsComponent)
    component = fixture.componentInstance
    fixture.detectChanges()
  })

  it('should create', () => {
    expect(component).toBeTruthy()
  })
})
