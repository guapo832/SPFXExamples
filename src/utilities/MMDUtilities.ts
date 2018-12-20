import {IManagedMetadata} from '../model/IManagedMetadata';
import {IHashTable} from '../model/IHashTable';
export class MMDUtilities{

  public static buildHash(values:IManagedMetadata[]):IHashTable<IManagedMetadata>{
    if(values === undefined || values == null) return null;
    var result:IHashTable<IManagedMetadata> = values.reduce(function(map, obj) {
      map[obj.TermGuid] = obj;
      return map;
  }, {});
  return result;
  }

  public static getTermsFromTermSetKeyword(termSetId:string,keyword:string):Promise<IManagedMetadata[]> {
    return new Promise((resolve,reject) => {
      try{
            if(!termSetId || termSetId ==="")
            {
              reject(new Error("No termset Id specified"));
              return;
            }

            if(!keyword || keyword ===""){
               resolve([]);
               return;
            }
            
          let mmdItems: IManagedMetadata[];
          let Terms: SP.Taxonomy.Term[] = new Array();
          var context: SP.ClientContext =  SP.ClientContext.get_current();
        
        
          let taxSession:SP.Taxonomy.TaxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
         let termStore:SP.Taxonomy.TermStore = taxSession.getDefaultSiteCollectionTermStore();

          let tset:SP.Taxonomy.TermSet =termStore.getTermSet(new SP.Guid(termSetId));
          let  match:SP.Taxonomy.LabelMatchInformation = new SP.Taxonomy.LabelMatchInformation(context);
          match.set_defaultLabelOnly(true);
          match.set_resultCollectionSize(10);
          match.set_stringMatchOption(SP.Taxonomy.StringMatchOption.startsWith);
          match.set_trimUnavailable(true);
          //match.set_trimDeprecated(true);
          match.set_termLabel(keyword);
          let terms:SP.Taxonomy.TermCollection  = tset.getTerms(match);
          
          
          //= termStore.getTermSet(new SP.Guid(termSetId)).get_terms();
          
          
      
          context.load(terms, 'Include(Labels, TermsCount, CustomSortOrder, Id, Name, PathOfTerm, TermSet.Name)');
          
                 
            context.executeQueryAsync( (sender,args) => {
              var termItems= [];
              
              var termEnumerator = terms.getEnumerator();
            while (termEnumerator.moveNext()) {
                     var currentTerm = termEnumerator.get_current();
                     termItems.push(currentTerm);
                }
              mmdItems = termItems.map((element):IManagedMetadata =>{
                var itm = element;
                let id = element.get_id().toString();
                let label = element.getDefaultLabel(1033).get_value();
                let name = element.get_name();
                let path = element.get_pathOfTerm();
                let TermsCount = element.TermsCount;
                let CustomSortOrder = element.CustomSortOrder;
                return({DisplayName:name, Label:label, TermGuid:id, PathofTerm:path});
              
            });
            console.log(mmdItems);
            resolve(mmdItems);
            },(sender,args)=>{
              let error = new Error();
              error.message = "Unable to retreive term from termId: " + termSetId;
              console.log(error.message);
              resolve([]);
            });
          } catch(e){
            console.log(e);
            reject(e);
          }
             
         
          
          
          });
  }

  public static getTermItemsFromTermSet(termId: string): Promise<IManagedMetadata[]>{
    return new Promise((resolve,reject) => {
      try{
            if(!termId)
            {
              debugger;
              reject(new Error("No termId specified"));
            }
            
          let mmdItems: IManagedMetadata[] = [];
          let Terms: SP.Taxonomy.Term[] = new Array();
          var context: SP.ClientContext =  SP.ClientContext.get_current();
        
        
          var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
          var termStore = taxSession.getDefaultSiteCollectionTermStore();
          let terms = termStore.getTerm(new SP.Guid(termId)).get_terms();

          context.load(terms, 'Include(Labels, TermsCount, CustomSortOrder, Id, Name, PathOfTerm, TermSet.Name)');
            context.executeQueryAsync( (sender,args) => {
              var termItems = [];
              var termEnumerator = terms.getEnumerator();  
            while (termEnumerator.moveNext()) {
                     let currentTerm:SP.Taxonomy.Term = termEnumerator.get_current();
                     let id: string = currentTerm.get_id().toString();
                     let label: string = currentTerm.getDefaultLabel(1033).get_value();
                     let name:  string = currentTerm.get_name();
                     let path:string = currentTerm.get_pathOfTerm();
                     let childcount = currentTerm.get_termsCount();
                     mmdItems.push({ChildCount:childcount,DisplayName:name,Expanded:false,Label:label,PathofTerm:path,TermGuid:id});
                }

            console.log(mmdItems);
            resolve(mmdItems);
            },(sender,args)=>{
              let error = new Error();
              error.message = "Unable to retreive term from termId: " + termId;
              console.log(error.message);
              reject(error);
            });
          } catch(e){
            console.log(e);
            reject(e);
          }
             
         
          
          
          });
          
  }

 

    public static getTermItems(termSetId: string): Promise<IManagedMetadata[]> {
        return new Promise((resolve,reject) => {
    try{
          if(!termSetId)
          {
            reject(new Error("required Termset Id unspecified in webpart properties"));
          }
          
        let mmdItems: IManagedMetadata[] = [];
        let Terms: SP.Taxonomy.Term[] = new Array();
        var context: SP.ClientContext =  SP.ClientContext.get_current();
      
      
        var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        var termStore = taxSession.getDefaultSiteCollectionTermStore();
        
        //var group = termStore.getGroup(new SP.Guid(groupId));
        //let terms = group.get_termSets().getByName(termSetName).getAllTerms();
        let terms = termStore.getTermSet(new SP.Guid(termSetId)).get_terms();
        
    
        context.load(terms, 'Include(Labels, TermsCount, CustomSortOrder, Id, Name, PathOfTerm, TermSet.Name)');
        
               
          context.executeQueryAsync( (sender,args) => {
            var termItems = [];
            var termEnumerator = terms.getEnumerator();
            while (termEnumerator.moveNext()) {
              let currentTerm:SP.Taxonomy.Term = termEnumerator.get_current();
              let id: string = currentTerm.get_id().toString();
              let label: string = currentTerm.getDefaultLabel(1033).get_value();
              let name:  string = currentTerm.get_name();
              let path:string = currentTerm.get_pathOfTerm();
              let childcount = currentTerm.get_termsCount();
              mmdItems.push({ChildCount:childcount,DisplayName:name,Expanded:false,Label:label,PathofTerm:path,TermGuid:id});
         }
          resolve(mmdItems);
          },(sender,args)=>{
            console.log("unable to retreive Termset: " + termSetId);
            let error = new Error();
            error.message = "Unable to retreive Termset: " + termSetId;
            reject(error);
          });
        } catch(e){
          console.log(e);
          reject(e);
        }
           
        //});
        
        
        });
        
        
    
              
        
      
    
      }
    
}
