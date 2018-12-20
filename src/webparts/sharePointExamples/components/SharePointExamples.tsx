import * as React from 'react';
import styles from './SharePointExamples.module.scss';
import { ISharePointExamplesProps } from './ISharePointExamplesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPUtilities } from '../../../utilities/SPUtilities';
import { TaxonomyPicker , TaxonomyPickerFull} from '../../../components/TaxonomyPicker';

export interface ISharePointExamplesState {
  relativeweburl?:string;
  absoluteweburl?:string;
}


export default class SharePointExamples extends React.Component<ISharePointExamplesProps, ISharePointExamplesState> {

  constructor(props){
    super(props);
   this.state = {relativeweburl:"test",absoluteweburl:"test"};
  }



  protected componentDidMount(){
      SPUtilities.loadCSOM()
      .then(()=>{
        let ctx = SP.ClientContext.get_current();
        let web = ctx.get_web();
        ctx.load(web);
        ctx.executeQueryAsync(
          (sender, args)=>{
            this.setState({relativeweburl:web.get_serverRelativeUrl(), absoluteweburl:web.get_url()});
          },
          (sender,args) =>{console.error(args.get_stackTrace());}
        );
      })
      .catch(error => console.error(error));
  }

  public render(): React.ReactElement<ISharePointExamplesProps> {
    return (

      
      <div className={ styles.sharePointExamples }>
      <h1>
        Hello from CSOM web.get_url() -> : {this.state.absoluteweburl}<br />
        Hello from CSOM web.get_serverRelativeUrl() -> : {this.state.relativeweburl}
      </h1>
        <div>
          <div>
            <h2>Components:</h2>
                MMD: <TaxonomyPicker choices={[{DisplayName:"test",Label:"test",TermGuid:"333"}]} onChange={()=>{}}></TaxonomyPicker><br />

                MMDFull Audit Area Process Sub Process: <TaxonomyPickerFull onChange={()=>{}} TermsetId="1d0b9e44-c6f7-44e3-826c-d3a5611f3a0d"></TaxonomyPickerFull>
      <br />
            </div>
          </div>
        </div>
      
    );
  }
}
