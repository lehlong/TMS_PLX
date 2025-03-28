import {BaseFilter} from '../base.model';

export class AccountFilter extends BaseFilter {
  GroupId : string = '';
  IsDescending: boolean = true;
  SortColumn: string = 'name';
  PartnerId! : number ;
  AccountType : string = '';
  IsActive! : boolean;
}


export interface AdAccount {
    userName: string;
    fullName: string;
    groupId: string;
    isActive: boolean;
    accountGroup: AdAccountGroup;
  }
  export interface AdAccountCreate {
    userName: string;
    fullName: string;
    companyCode: string;
    position: string;
    department: string;
    email: string;
    phoneNumber: string;
    isActive: boolean;
  }
  
  export interface AdAccountGroup {
    id: string;
    name: string;
    isActive: boolean;
    notes: string;
    listAccount: AdAccount[];
  }
  export interface AdAccountUpdateInfor {
    userName: string;
    fullName: string;
    phoneNumber: string;
    email: string;
    address: string;
  }
  export interface AdChangePassword {
    userName: string;
    oldPassword: string;
    newPassword: string;
  }
  
  export interface AdAccountModel {
    userName?: string;
    fullName?: string;
    groupId?: string;
    isActive?: boolean | string;
  }
  