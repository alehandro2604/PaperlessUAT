//these are the fields inside the eForm list
export interface EForm {
id:number;
title:string;
prefix:string;
url:string;
taskListName:string;
taskListGuid:string;
docLibraryName:string;
isHRWorkflow:boolean;
isSAPWorkflow:boolean;
HROrder:number;
SAPOrder:number;
/** Numeric order for the All Files library dropdown. */
AllFiles:number;
/** When true, the form's DocLibraryName appears in the All Files dropdown. */
AllFilesCheckBox:boolean;
}