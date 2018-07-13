import * as React from 'react'

import { DocumentCard, DocumentCardTitle, DocumentCardPreview, IDocumentCardPreviewProps, DocumentCardActivity } from 'office-ui-fabric-react';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import { GraphDriveItem } from './Entities';
import GraphClient from './GraphClient';

interface IDocComponentProps {
    driveItem: GraphDriveItem
}

interface IDocComponentState {
    driveItem: GraphDriveItem
}

export default class DocumentComponent extends React.Component<IDocComponentProps, IDocComponentState> {

    private _getFullItem(): void {
        let gClient: GraphClient = new GraphClient();
        console.debug('Trying to get the full details for this item');
        gClient.graphRequest(
            "https://graph.microsoft.com/beta/me/insights/trending/" + this.state.driveItem.docId + "/resource",
            (doc) => {
                let docAsJson = JSON.parse(doc);
                let tmp: GraphDriveItem = {
                    docId: this.state.driveItem.docId,
                    title: this.state.driveItem.title,
                    author: docAsJson.lastModifiedBy.user.displayName
                }
                this.setState({
                    driveItem: tmp
                })
            },
            (error) => {
                console.error(`Error retrieving item details for item ${this.state.driveItem.title}. Continuing.`);
            }
        )
    }

    public componentDidUpdate(prevProps, prevState, snapshot) {
        if(prevProps != this.props && this.props.driveItem){
            this.setState({
                driveItem: this.props.driveItem
            })
            //this._getFullItem();
        }
    }

    public componentWillMount() {
        this.setState({
            driveItem: this.props.driveItem
        })
    }

    public componentDidMount() {
        //this._getFullItem();
    }

    public render() {
        console.debug('Rendering drive item');

        const prevProps: IDocumentCardPreviewProps = {
            previewImages:[
                {
                    name: this.state.driveItem.title,
                    url: this.state.driveItem.webUrl,
                    previewImageSrc: this.state.driveItem.previewImageUrl,
                    imageFit: ImageFit.cover
                }
            ]
        }
        
        return (
                
                <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg4 teamsInsightCard" key={this.state.driveItem.docId}>
                    <DocumentCard onClickHref={this.state.driveItem.webUrl}>
                        <div className="ms-ConversationTile-TitlePreviewArea">
                            <DocumentCardPreview {...prevProps} />
                            <DocumentCardTitle title={this.state.driveItem.title} shouldTruncate={false}/>
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
    }
}