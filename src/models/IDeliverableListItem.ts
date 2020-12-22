export interface IDeliverableListItem {
    ID:number;
    Title:string;
    DType?:string;
    LOB:{
        Title:string;
        Id:number;
    };
    Scenario:string;
    Status:string;
    Description1:string;


}