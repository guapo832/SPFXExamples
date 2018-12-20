/*Private Methods */
import { Environment, EnvironmentType} from "@microsoft/sp-core-library";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { sp, CamlQuery } from "sp-pnp-js";

export class SPUtilities{
    public static FormatDate = (date): string => {
        if(date===undefined || date== null) {return "";}
        let newdate = new Date(date);
        return (newdate.getMonth() +1) + "/" + newdate.getDate() + "/" + (newdate.getFullYear());
    }

    private static _checkSPLoaded():boolean{
        if(SP === undefined ||  SP.ClientContext === undefined) {return false;}
        return true;
    }

   
    

    public static loadCSOM(): Promise<{}>{
                    let globalExportsName:string= "$_global_init";
                    var promise = new Promise((resolve,reject)=>{
                        let p:any = (window[globalExportsName]?
                            Promise.resolve(window[globalExportsName]):
                            SPComponentLoader.loadScript("/_layouts/15/init.js", { globalExportsName }));
                            p.catch((error) => { console.error(error); })
                            .then(($_global_init): Promise<any> => {
                                globalExportsName = "Sys";
                                p = (window[globalExportsName] ?
                                    Promise.resolve(window[globalExportsName]):
                                    SPComponentLoader.loadScript("/_layouts/15/MicrosoftAjax.js", { globalExportsName }));
                                return p;
                              }).catch((error) => {
                                  console.error(error);
                               })
                               .then((Sys): Promise<any> => {
                                globalExportsName = 'SP';
                                p = ((window[globalExportsName] && window[globalExportsName].ClientRuntimeContext)?
                                 Promise.resolve(window[globalExportsName]):
                                 SPComponentLoader.loadScript("/_layouts/15/SP.Runtime.js", { globalExportsName }));
                                return p;
                              }).catch((error) => { 
                                  console.error(error);
                              })
                              .then((SP): Promise<any> => {
                                globalExportsName = "SP";
                                 p = ((window[globalExportsName] && window[globalExportsName].ClientContext)?
                                 Promise.resolve(window[globalExportsName]):
                                 SPComponentLoader.loadScript("/_layouts/15/SP.js", { globalExportsName }));
                                return p;
                              }).catch((error) => {
                                  console.error(error);
                               }).then((SP)=>{
                                globalExportsName = "SP.Taxonomy";
                                p = ((window[globalExportsName])?
                                Promise.resolve(window[globalExportsName]):
                                SPComponentLoader.loadScript( "/_layouts/15/SP.taxonomy.js", {
                                    globalExportsName: "SP.Taxonomy"
                                }));
                               return p;
                               }).catch((error) => {
                                console.error(error);
                            })
                            .then((SP)=> {resolve(SP);} );
                    });
                    return promise;
            }

        public static getTermFieldLabel(internalFidleName:string,itmId:number, listtitle:string):Promise<string> {
        let querystring:string = "<View>" +
        "<Query>" +
        "<Where>" +
        "<Eq><FieldRef Name='ID'/><Value Type='Number'>" + itmId + "</Value></Eq>" +
        "</Where></Query>" +
        "</View>";
        return new Promise((resolve,reject)=> {
            let query:CamlQuery = {
                ViewXml: querystring
            };
            sp.web.lists.getByTitle(listtitle).getItemsByCAMLQuery(query).then((items)=>{
                if(items.length===0){
                   let error = new Error();
                    error.message = "No list item found with ID: " + itmId + " in list: " + listtitle;
                    reject(error);
                }
                resolve(items[0][internalFidleName] && items[0][internalFidleName].Label);
            })
            .catch((error)=> {
               reject(error);
            });
        });
    }
            
        }
