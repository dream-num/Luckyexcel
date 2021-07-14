import { IluckysheetDataVerificationType } from "./ToLuckySheet/ILuck";


export interface IuploadfileList { 
    [index:string]:string 
} 

export interface stringToNum {
    [index:string] : number
}

export interface numTostring {
    [index:number] : string
}

export interface IattributeList {
    [index:string]:string
}

export interface IDataVerificationMap {
    [key: string]: IluckysheetDataVerificationType;
}

export interface IDataVerificationType2Map {
    [key: string]: { [key: string]: string };
}
