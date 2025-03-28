import { BaseFilter } from "../base.model"

export class TemplateListTablesGroupsFilter extends BaseFilter {
    code: string = ''
    id: string = ''
    name: string = ''
    isActive?: boolean | string | null
    SortColumn: string = ''
    IsDescending: boolean = true
}