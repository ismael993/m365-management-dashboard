import { ComponentFixture, TestBed } from '@angular/core/testing';

import { OfficeAccounts } from './office-accounts';

describe('OfficeAccounts', () => {
  let component: OfficeAccounts;
  let fixture: ComponentFixture<OfficeAccounts>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [OfficeAccounts],
    }).compileComponents();

    fixture = TestBed.createComponent(OfficeAccounts);
    component = fixture.componentInstance;
    await fixture.whenStable();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
