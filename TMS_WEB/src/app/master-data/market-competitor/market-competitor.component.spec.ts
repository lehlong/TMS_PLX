import { ComponentFixture, TestBed } from '@angular/core/testing';

import { MarketCompetitorComponent } from './market-competitor.component';

describe('MarketCompetitorComponent', () => {
  let component: MarketCompetitorComponent;
  let fixture: ComponentFixture<MarketCompetitorComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [MarketCompetitorComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(MarketCompetitorComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
