import { Component } from '@angular/core';
import { OfficeAccounts } from './office-accounts/office-accounts';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [OfficeAccounts],
  template: `<app-office-accounts></app-office-accounts>`
})
export class App {}
