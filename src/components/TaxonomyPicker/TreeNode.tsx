import * as React from "react";
import {SPUtilities} from "../../utilities/SPUtilities";
import {IManagedMetadata} from "../../model/IManagedMetadata";
import { MMDUtilities} from "../../utilities/MMDUtilities";
import {  cloneDeep} from "@microsoft/sp-lodash-subset";
import { Dialog, DialogFooter } from "office-ui-fabric-react/lib-amd/Dialog";
import {PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib-amd/Button";

import { IHashTable } from "../../model/IHashTable";

export interface ITreeNodeProps {
    term?:IManagedMetadata;
    expanded?:boolean,
    index?:number;
    onToggleNode?:(terGuid:string,index:number) =>void;
    children?:IManagedMetadata[];
    focusedTerm?:IManagedMetadata;
    onFocusTerm:(term:IManagedMetadata) =>void;

    
}

export interface ITreeNodeState{
    termNodesJsx?:JSX.Element[];
    children?:IManagedMetadata[];
    tnnClassName?:string;
    focusedTerm?:IManagedMetadata;
}

export class  ForesightMMDTreeNode extends React.Component<ITreeNodeProps,ITreeNodeState>
{
    constructor(props: ITreeNodeProps) {
        super(props);
        let children = this.props.children !==undefined?cloneDeep(this.props.children):[];
       this._onClickTwisty = this._onClickTwisty.bind(this);
       this.state={termNodesJsx:[],children:children,tnnClassName:"tnn",focusedTerm:null}
        this._getTerms = this._getTerms.bind(this);
        this._nodeBlur = this._nodeBlur.bind(this);
        this._onNodeClick = this._onNodeClick.bind(this);
        this._onNodeDoubleClick = this._onNodeDoubleClick.bind(this);
        this._onNodeHover = this._onNodeHover.bind(this);
}

protected componentDidMount(){
    if(this.props.expanded) {
         SPUtilities.loadCSOM()
        .then(this._getTermSet.bind(this))
        .catch(this._handleError.bind(this))
        .then((termset:IManagedMetadata[])=> {
        this.termsHash = MMDUtilities.buildHash(termset);
        this.setState({children:termset, tnnClassName:"tnn",focusedTerm:null},()=>{
        
            this._buildTermNodeJsx();

        });
        });
    }
    
}

    private termsHash:IHashTable<IManagedMetadata>;

protected componentDidUpdate(prevProps:ITreeNodeProps){
    if((prevProps.expanded === false && this.props.expanded === true && this.props.children)
    || (prevProps.focusedTerm !== this.props.focusedTerm)){
        this.termsHash = MMDUtilities.buildHash(this.props.children);
        let selectedCSS =this.props.focusedTerm && this.props.focusedTerm.TermGuid === this.props.term.TermGuid?"treenodeonfocus":"tnn";
        this.setState({children:this.props.children,tnnClassName:selectedCSS,focusedTerm:this.props.focusedTerm},this._buildTermNodeJsx);
     }
}

private _getTerms(termGuid:string,index:number){
    const { children} = this.state;
    children[index].Expanded = !children[index].Expanded;
    if(children[index].Children === undefined || children[index].Children === null){
        MMDUtilities.getTermItemsFromTermSet(termGuid).then(
            (child) =>{
                this.termsHash[termGuid].Children = child;
                let childrenArr: IManagedMetadata[] = Object.keys(this.termsHash).map((key):IManagedMetadata=>{
                    return this.termsHash[key];
                })
                this.setState({children:childrenArr},()=>{
                    this._buildTermNodeJsx();
                });
            }
        )
        } else {
            this._buildTermNodeJsx();
        }
}


public render() {
    const { term } = this.props;
        return(<li>
             <div className="treenodediv">
             <span className="_ImgContainer">{term.ChildCount > 0?
                 <a href="" onClick={this._onClickTwisty}>
                 <img className="ti"
                 src={this.props.expanded !== undefined &&
                  this.props.expanded === true?
                  "/_layouts/15/images/MDNExpanded.png":
                  "/_layouts/15/images/MDNCollapsed.png"}></img>
                 </a>:""}         
             </span>
             <img width={16} height={16} src="/_layouts/15/Images/EMMTerm.png" />
             <span 
             onClickCapture={this._onNodeClick} 
             className={this.state.tnnClassName}
             onMouseOverCapture={this._onNodeHover}
             onMouseOutCapture={this._nodeBlur}
             onDoubleClick={this._onNodeDoubleClick}
             >
             <span className="ms-input">
                 <span className="ms-pagetitle" style={{minWidth:"100px"}}>
                     {this.props.term.DisplayName}
                 </span>
             </span></span>
            </div>
            <ul style={{display:this.props.expanded?"block":"none"}}>
                {this.state.termNodesJsx}
            </ul>
        </li>);

}

private _onNodeClick(event){
    this.setState({focusedTerm:this.props.term},()=>{this.props.onFocusTerm(this.props.term);});
  
}

private _onNodeDoubleClick(event){
 this.props.onFocusTerm(this.props.term);
}

private _onNodeHover(event){
if(this.state.tnnClassName !== "treenodeonfocus"){
    this.setState({tnnClassName:"treenodehover"});
}
}

private _handleOnFocused(term:IManagedMetadata){
       
        this.props.onFocusTerm(term);

}



private _nodeBlur(event){
    if(this.state.tnnClassName === "treenodehover"){
        this.setState({tnnClassName:"tnn"});
    }

}
private _buildTermNodeJsx(){
    if(this.state.children){
    let nodes:JSX.Element[] = this.state.children.map((itm,idx)=>{
        let children = this.state.children[idx].Children !== null?this.state.children[idx].Children:[];
        let expanded = this.state.children[idx].Expanded || false;
        return<ForesightMMDTreeNode
        focusedTerm={this.state.focusedTerm}
        onFocusTerm = {this._handleOnFocused.bind(this)}
       expanded = {expanded}
        onToggleNode={this._getTerms}
        index={idx}
        children ={children}
        term={itm}></ForesightMMDTreeNode>})
        this.setState({termNodesJsx:nodes});
    }
}

/*
private _buildTermNodeJsx(){
    const {termNodesJsx} = this.state;
    let rtnval = [];
    if(this.state.children !== undefined && this.state.children !== null){
        for(let i = 0 ; i<this.state.children.length; i++){
               let expanded:boolean = termNodesJsx[i].props.expanded || false;
                rtnval.push(<ForesightMMDTreeNode
                expanded={expanded}
                onToggleNode={this._getTerms}
                index={i}
                term={this.state.children[i]}></ForesightMMDTreeNode>);
            
        }
        this.setState({termNodesJsx:rtnval});
    }
}
*/
private _onClickTwisty(e){
    e.preventDefault();
    this.props.onToggleNode(this.props.term.TermGuid,this.props.index);
}

private _getTermSet():Promise<IManagedMetadata[]> {
    return MMDUtilities.getTermItemsFromTermSet(this.props.term.TermGuid);
}

private _handleError(error:Error){
    console.error(error);
}



}