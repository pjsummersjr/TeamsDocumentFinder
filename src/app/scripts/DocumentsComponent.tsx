import * as React from 'react';

export interface SPODocument {
    docId: string;
    title: string;
    author?: string;
    webUrl?: string;
    previewImageUrl?: string;
}

import { DocumentCard, DocumentCardTitle, DocumentCardPreview, IDocumentCardPreviewProps, DocumentCardActivity } from 'office-ui-fabric-react';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';

export interface IDocumentsComponentState  {
    documents?: SPODocument[];
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

    private _processDocs(rawData:string): SPODocument[] {
        let tmpDocs: SPODocument[] = new Array();
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
        console.debug('Rendering documents')
        if(this.state.documents && this.state.documents.length > 0) {
            console.debug('Documents array not null or empty');
            
            return (this.state.documents.map((item, index) => {
                    const prevProps: IDocumentCardPreviewProps = {
                        previewImages:[
                            {
                                name: item.title,
                                url: item.webUrl,
                                previewImageSrc: item.previewImageUrl,
                                imageFit: ImageFit.cover
                            }
                        ]
                    }
                    
                    return (
                            
                            <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg4 teamsInsightCard" key={item.docId}>
                                <DocumentCard onClickHref={item.webUrl}>
                                    <div className="ms-ConversationTile-TitlePreviewArea">
                                        <DocumentCardPreview {...prevProps} />
                                        <DocumentCardTitle title={item.title} shouldTruncate={false}/>
                                        <DocumentCardTitle 
                                            title="This is the document summary of the document which is extract of key pieces of content from the content." 
                                            shouldTruncate={false}
                                            showAsSecondaryTitle={true}
                                            />
                                    </div>
                                    <DocumentCardActivity activity="Last modified June 28, 2018" people={[
                                        {name: 'Roko Kolar', profileImageSrc: '', initials: 'JH'}
                                    ]} />
                                </DocumentCard>
                            </div>
                        );
                })
            );
        }
        else {
            return (<div>No documents found</div>);
        }
    }
}