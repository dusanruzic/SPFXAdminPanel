export interface IList {
    Id: string;
    title: string;
    [index: string]: any;
}

export interface IListCollection {
    value: IList[];
}