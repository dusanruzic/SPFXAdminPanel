export interface IListItem {
    Id: number;
    Title: string;
    [index: string]: any;
}

export interface IListItemCollection {
    value: IListItem[];
}