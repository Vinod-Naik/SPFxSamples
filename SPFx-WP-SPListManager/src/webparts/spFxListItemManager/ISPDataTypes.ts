export interface ISPList {
    Id : string,
    Title ?: string,
    Template ?: string,
    Description ?: string
}

export interface ISPLists{
    value : ISPList[]
}
export interface ISPListItem {
    Id : string,
    Title ?: string,
    Modified ?: string
}

export interface ISPListItems{
    value : ISPListItem[]
}