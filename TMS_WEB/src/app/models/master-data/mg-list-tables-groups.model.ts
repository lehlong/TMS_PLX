import { BaseFilter } from "../base.model"

export class MgListTablesGroupsFilter extends BaseFilter {
    code: string = ''
    id: string = ''
    name: string = ''
    isActive?: boolean | string | null
    SortColumn: string = ''
    IsDescending: boolean = true
}