import * as React from "react";
import styles from "./TaxonomyPicker.module.scss";
import {IManagedMetadata} from "../../model/IManagedMetadata";
import { SPUtilities } from "../../utilities/SPUtilities";
import { MMDUtilities } from "../../utilities/MMDUtilities";
import { TagPicker, ITag } from "office-ui-fabric-react/lib/components/pickers/TagPicker/TagPicker";
import { TermsTreeDialog } from "./TermsTreeDialog";
import cloneDeep = require("lodash/cloneDeep");
declare var $: any;

export interface ITermsDropDownProps {
    choices?: IManagedMetadata[];
    onChange: (value: IManagedMetadata) => void;
    selectedValue?: IManagedMetadata;
    // allowMultiple?:boolean; //coming soon
}

export interface ITermsDropDownState {
}

export class  TaxonomyPicker extends React.Component<ITermsDropDownProps,ITermsDropDownState> {

    constructor(props: ITermsDropDownProps) {
        super(props);
        this._onTermChanged = this._onTermChanged.bind(this);
       
    }

    public render() {
        const {selectedValue,choices} = this.props;

        var options: JSX.Element[] = choices.map((term:IManagedMetadata): JSX.Element=>{
            let isSelected: boolean = (selectedValue && selectedValue.TermGuid == term.TermGuid) || false
            return <option value = {term.TermGuid} selected = {isSelected}>{term.DisplayName}</option>
        });

       return <span>
           <select onChange = {this._onTermChanged}>
       {options}</select></span>;
    }

    private _onTermChanged(event: React.FormEvent<HTMLSelectElement> ) {
        var selectedTerm:IManagedMetadata;
const {choices} = this.props;
        for (let term in choices){
            if (choices[term].TermGuid == event.currentTarget.value){
                selectedTerm = choices[term];
            }
        }
        this.props.onChange(selectedTerm);
    }

}


/* The following picker is not complete. */
/*  it will be a full treeview mmd picker */

export interface ITermPickerProps {
    TermsetId:string
    onChange: (value: IManagedMetadata[]) => void;
    selectedValue?: IManagedMetadata[];
}

export interface ITermPickerState {
    data?:ITag[];
    errorMessage?:string;
    pickerDialog?:boolean;
    pickerKey?:number
    
}



export class TaxonomyPickerFull extends React.Component<ITermPickerProps,ITermPickerState>{

    

    constructor(props) {
        super(props);
        this._onTermChanged = this._onTermChanged.bind(this);
       
        this.state = {
            errorMessage:null,
            data:[],
            pickerDialog:false,
            pickerKey:0
        }
        this._cancelDialog = this._cancelDialog.bind(this);
        this._saveDialog = this._saveDialog.bind(this);
    }

    protected componentDidMount(){
     
        SPUtilities.loadCSOM()
        .then(this._getTermSet.bind(this))
        .catch(this._handleError.bind(this))
        .then((termset:IManagedMetadata[])=> {
            
        });
        if(this.props.selectedValue && this.props.selectedValue.length>=0){
            const data:ITag[] = this.props.selectedValue.map((itm)=>{
                return {name:itm.PathofTerm,key:itm.TermGuid};
            });
            
            this.setState({data:data},()=>{
                
            });

        }
    }

    private _getTextFromItem(Item:ITag){
        return Item.name;
    }
    
    public render() {
          
        const {errorMessage} = this.state;

        if(errorMessage) {return <div style={{color:"red",fontStyle:"italic"}}>{errorMessage}</div>;}
       
        return <div className={styles.taxonomymmdpicker}>
        <TermsTreeDialog 
        onItemSelected={this._onItemSelected.bind(this)}
        TermsetId={this.props.TermsetId}
        isOpen={this.state.pickerDialog}
        onCancel={()=>{this.setState({pickerDialog:false})}}></TermsTreeDialog>
        
            <table style={{width:"100%"}}><tr><td>
            <TagPicker
                 getTextFromItem = {this._getTextFromItem.bind(this)}
                 key ={this.state.pickerKey}
                 defaultSelectedItems = {this.state.data}
                 onResolveSuggestions={this._onFilterChanged}
                 onChange = {this._onTermChanged.bind(this)}
                 pickerSuggestionsProps={{
                 suggestionsHeaderText: "Suggested Terms",
                 noResultsFoundText: "No Terms Found"
                 
               }}
               disabled ={(this.state.data.length ===1)}
             />
                </td><td><div className={styles.TaxPickerButton} onClick={(e)=>{this.setState({pickerDialog:true});}}></div></td></tr></table>
        </div>;
    }

    private _onItemSelected(data:IManagedMetadata[]){
        let newKey = this.state.pickerKey + 1;
        let tagData:ITag[] = data.map((itm)=>{ return {key:itm.TermGuid, name:itm.PathofTerm}});
        this.setState({data:tagData, pickerKey:newKey, pickerDialog:false},()=>{
            this.props.onChange(data);
        });
        }
   
    private _onTermChanged(itms:ITag[]) {
       this.setState({data:itms});
    }

    private _saveDialog(){

    }

    private _cancelDialog(){

        this.setState({pickerDialog:false});

    }
    
      private _onFilterChanged = (filterText: string, termList: ITag[]):Promise<ITag[]>  => {

        return new Promise((resolve,reject) =>{
            MMDUtilities.getTermsFromTermSetKeyword(this.props.TermsetId,filterText)
            .then((terms:IManagedMetadata[])=>{
                let suggestions:ITag[] = terms.map((itm) =>{
                    return {key:itm.TermGuid, name:itm.PathofTerm.replace(new RegExp(";", "g"),":")};
                });
                resolve(suggestions);
            }).catch((error)=> {
                console.error(error.message);
                resolve([]);
            })
            .then(()=>{resolve([])});
        });
        
      };







    private _getTermSet():Promise<IManagedMetadata[]> {
        return MMDUtilities.getTermItems(this.props.TermsetId);
    }

    private _handleError(error:Error){
        console.error(error);
        this.setState({errorMessage:error.message});
    }

    
}