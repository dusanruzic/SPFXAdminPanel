export interface IListField {
    Id: string;
    Title: string;
    Internalname: string;
    TypeAsString: string;
    [index: string]: any;
}

export interface IListFieldCollection {
    value: IListField[];
}