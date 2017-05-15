export interface ISPList {
    Id : string,
    Title ?: string,
    ItemCount ?: number,
    Template ?: string,
    Description ?: string
}

export interface ISPLists{
    value : ISPList[]
}