import * as React from "react";
import {SPUtilities} from "../../utilities/SPUtilities";
import {IManagedMetadata} from "../../model/IManagedMetadata";
import {IHashTable} from "../../model/IHashTable";
import { ClientSideWebPartManager } from "@microsoft/sp-webpart-base"
import { MMDUtilities} from "../../utilities/MMDUtilities";
import { Dialog, DialogFooter, IDialog } from "office-ui-fabric-react/lib-amd/Dialog";
import {PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib-amd/Button";
import { ForesightMMDTreeNode } from "./TreeNode";
import styles from "./TermsTreeDialog.module.scss";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { TagPicker, ITag, BasePicker, IBasePickerProps } from "office-ui-fabric-react/lib/components/pickers";
import cloneDeep = require("lodash/cloneDeep");

export interface ITermsPickerDialogProps {
    isOpen:boolean
    TermsetId:string;
    onSave?:(data:IManagedMetadata[]) => void;
    onCancel?:() => void;
    onItemSelected: (data:IManagedMetadata[]) => void;
}

export interface ITermsPickerDialogState{
    terms?:IManagedMetadata[];
    termNodesJsx?:JSX.Element[];
    focusedTerm?:IManagedMetadata;
    pickerKey?:number;
    selectedTerms?:IManagedMetadata[];
}



export class  TermsTreeDialog extends React.Component<ITermsPickerDialogProps,ITermsPickerDialogState>
{

    private weburl = "";
private termsHash: IHashTable<IManagedMetadata>;
    constructor(props: ITermsPickerDialogProps) {
        super(props);
       this.state = {terms:[],pickerKey:0};
       this.termsHash= {};
        this._cancelDialog = this._cancelDialog.bind(this);
        this._saveDialog = this._saveDialog.bind(this);
        this._getTerms = this._getTerms.bind(this);
        
    }

    protected componentDidMount(){
      
        SPUtilities.loadCSOM()
        .then(this._getTermSet.bind(this))
        .catch(this._handleError.bind(this))
        .then((termset:IManagedMetadata[])=> {
           this.termsHash = MMDUtilities.buildHash(termset);
           let ctx = SP.ClientContext.get_current();
           let web = ctx.get_web();
           ctx.load(web);
           ctx.executeQueryAsync((sender,obj)=>{
                this.weburl = web.get_serverRelativeUrl();
                SPComponentLoader.loadCss(this.weburl + "/_layouts/15/1033/styles/termstoremanager.css");
                this.setState({terms:termset},this._buildTermNodeJsx);
            
        },
        (sender,args)=>{
            console.error(args.get_stackTrace());
        });
          
        });
       
    }

    private _buildTermNodeJsx(){

        let nodes:JSX.Element[] = this.state.terms.map((itm,idx)=>{
            let children = this.state.terms[idx].Children !== null?this.state.terms[idx].Children:[];
            let expanded = this.state.terms[idx].Expanded || false;
            return<ForesightMMDTreeNode
            focusedTerm = {this.state.focusedTerm}
            onFocusTerm={this._handleOnFocused.bind(this)}
           expanded = {expanded}
            onToggleNode={this._getTerms}
            index={idx}
            children ={children}
            term={itm}></ForesightMMDTreeNode>})
            this.setState({termNodesJsx:nodes});
    }
    private dialog:IDialog;
    

public render(){
    let data =this.state.selectedTerms || [];
    let tags:ITag[] = data.map((itm)=>{return {key:itm.TermGuid, name:itm.PathofTerm}});
    return ( <Dialog 
        containerClassName={ "ms-dlgContent " + styles.TermsTreeDialog}
        isOpen={this.props.isOpen}
        isClickableOutsideFocusTrap={true}
        onDismiss={this._cancelDialog}>
        <div style={{height:"75%"}} >
        <table width="100%">
            <tbody>
                <tr>
                    <td style={{width:"100%", height:"0%"}}>
                        <table width="100%" className="ms-dialogHeader" style={{paddingTop:"8px", paddingBottom:"10px"}}>
                            <tbody>
                                <tr>
                                    <td style={{textAlign:"center", verticalAlign:"middle",paddingRight:"15px", paddingLeft:"15px"}}>
                                        <img alt="" src={this.weburl + "/_layouts/images/EMMDoubleTag.png"}/>
                                    </td>
                                    <td className="ms-dialogHeaderDescription" style={{width:"100%"}}>
                                        <div className="none-wordbreak"></div>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td className="headingDivider"></td>
                </tr>
                <tr>
                    <td className="bodyTableCell">
                        <div className={styles.dlgContainer}>
                        <table width="100%" height="100%">
                            <tbody>
                                <tr>
                                    <td className="ms-dialogBodyMain">
                                        <div>
                                        <div className="wt-treecontainer">
                                            <ul className={"TmtTree "} style={{height:"100%"}}>
                                            {this.state.termNodesJsx}
                                            
                                            </ul>
                                        </div>
                                       </div>
                                    </td>
                                </tr>                                
                            </tbody>
                        </table>
                        </div>
                    </td>
                </tr>
            </tbody>
            
        </table>
        <table width="100%">
            <tr>
                <td style={{width:"0%"}}>
                <PrimaryButton
        text="Select"
        onClickCapture={this._selectTerm.bind(this)}></PrimaryButton>
                </td>
                <td>
<TagPicker
defaultSelectedItems={tags}
onResolveSuggestions={this._onFilterChanged.bind(this)}
onChange = {this._onTermChanged.bind(this)}
key = {this.state.pickerKey}
pickerSuggestionsProps={{
suggestionsHeaderText: "Suggested Terms",
noResultsFoundText: "No Terms Found"
}}
disabled ={true}></TagPicker>

                </td>
            </tr>
        </table>
</div>
       
        <DialogFooter>
         <PrimaryButton onClick={this._saveDialog} text="Save" />
         <DefaultButton onClick={this._cancelDialog} text="Cancel" />
       </DialogFooter>
     </Dialog>);

}
private _onFilterChanged(){
    return new Promise((resolve,reject) =>{
        resolve([]);
    });
    
}
private _onTermChanged(){

}
   private _handleOnFocused(term:IManagedMetadata){
       this.setState({focusedTerm:term},this._buildTermNodeJsx);
   }
    private _getTerms(termGuid:string,index:number){
        const { termNodesJsx, terms} = this.state;
        terms[index].Expanded = !terms[index].Expanded;
        if(terms[index].Children === undefined || terms[index].Children === null){
            MMDUtilities.getTermItemsFromTermSet(termGuid).then(
                (children) =>{
                    this.termsHash[termGuid].Children = children;
                    let childrenArr: IManagedMetadata[] = Object.keys(this.termsHash).map((key):IManagedMetadata=>{
                        return this.termsHash[key];
                    })
                    this.setState({terms:childrenArr},this._buildTermNodeJsx);
                }
            )
            } else {
                this._buildTermNodeJsx();
            }
    }

    private _saveDialog(){
         let rtnval:IManagedMetadata[] = cloneDeep(this.state.selectedTerms) || [];
         this.setState({focusedTerm:null, selectedTerms:[]},()=>{
            this.props.onItemSelected(rtnval);
         });
    }

    private _cancelDialog(): void{
     this.setState({focusedTerm:null,selectedTerms:[]},() => {
         this.props.onCancel && this.props.onCancel();
     });
    }
    
    private _getTermSet():Promise<IManagedMetadata[]> {
        return MMDUtilities.getTermItems(this.props.TermsetId);
    }

    private _handleError(error:Error){
        console.error(error);
    }

    private _selectTerm(event){
        let data:IManagedMetadata[] = this.state.selectedTerms || [];
        let exists = false;
        for(let i=0; i<data.length; i++){
            if(data[i].TermGuid === this.state.focusedTerm.TermGuid)
            exists = true;
        }
        if(!exists) {
            data.push(this.state.focusedTerm)
        }
        this.setState({selectedTerms:data,pickerKey:(this.state.pickerKey+ 1)})
    }

}