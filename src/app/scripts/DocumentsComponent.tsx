import * as React from 'react';

import { GraphDriveItem } from './Entities';

import DocumentComponent from './DocumentComponent';

export interface IDocumentsComponentState  {
    documents?: GraphDriveItem[];
    isLoading?: boolean;
}

export interface IDocumentsComponentProps  {
    documentsData?:string;
}

export default class DocumentsComponent extends React.Component<IDocumentsComponentProps, IDocumentsComponentState> {

    constructor(props) {
        super(props);
        this.state = {
            documents:[],
            isLoading: false
        }
    }

    private _processDocs(rawData:string): GraphDriveItem[] {
        let tmpDocs: GraphDriveItem[] = new Array();
        console.debug('Processing raw document data');

        let docJson = JSON.parse(rawData);
        docJson.value.map((item) => {

            tmpDocs.push({
                docId: window.btoa(item.id),
                title: item.resourceVisualization.title,
                previewImageUrl: item.resourceVisualization.previewImageUrl,
                webUrl: item.resourceReference.webUrl
            });
        });
        return tmpDocs;
    }

    public componentDidUpdate(prevProps, prevState, snapshot) {
        if(prevProps != this.props && this.props.documentsData && this.props.documentsData.length > 0){
            this.setState({
                documents: this._processDocs(this.props.documentsData)
            })
        }
    }

    public componentDidMount() {
        if(this.props.documentsData && this.props.documentsData.length > 0){                           
            this.setState({
                documents: this._processDocs(this.props.documentsData),
                isLoading: false
            });
        }
    }

    public render() {
        
        if(this.state.documents && this.state.documents.length > 0) {
            console.debug('Documents array not null or empty');
            
            return (this.state.documents.map((item, index) => {
                    return (<DocumentComponent driveItem={item} />);
                })
            );
        }
        else {
            return (<div>No documents found</div>);
        }
    }
}