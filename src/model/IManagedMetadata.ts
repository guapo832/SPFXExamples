export interface IManagedMetadata{
    Parent?:IManagedMetadata;
    TermGuid:string;
    Label:string;
    DisplayName:string;
    ChildCount?:number;
    Children?:IManagedMetadata[];
    PathofTerm?:string;
    Expanded?:boolean;
}