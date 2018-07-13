import * as React from 'react';
import {
    PrimaryButton,
    TeamsComponentContext,
    ConnectedComponent,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface
} from 'msteams-ui-components-react';
import { render } from 'react-dom';
import { TeamsBaseComponent, ITeamsBaseComponentProps, ITeamsBaseComponentState } from './TeamsBaseComponent'
import GraphClient from './GraphClient';
import DocumentsComponent from './DocumentsComponent';

/**
 * State for the documentFinderTab React component
 */
export interface IDocumentFinderTabState extends ITeamsBaseComponentState {
    entityId?: string;
    displayName?: string;
    location?: string;
    jobTitle?: string;
    alerts?: string;
    documents?:string;
}

/**
 * Properties for the documentFinderTab React component
 */
export interface IDocumentFinderTabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the Document Finder content page
 */
export class DocumentFinderTab extends TeamsBaseComponent<IDocumentFinderTabProps, IDocumentFinderTabState> {
 
    private token: any = "";
    private currentDocumentEndPoint: any = "https://graph.microsoft.com/beta/me/insights/trending";
    constructor(props: IDocumentFinderTabProps, state: IDocumentFinderTabState) {
        super(props, state);
    }

    public error = (error: any) => {
        console.error("An error occurred. Trace appropriately");
    }

    public reloadContent = () => {
        const client: GraphClient = new GraphClient();
        client.graphRequest(this.currentDocumentEndPoint, 
            (response) => { this.setState({
                documents: response
            })},
            (error) => {this.error(error);}
        )
    }

    public trending = () => {
        this.currentDocumentEndPoint = "https://graph.microsoft.com/beta/me/insights/trending?$filter=ResourceReference/type eq 'microsoft.graph.driveItem'";
        this.reloadContent();
    }

    public recent = () => {
        alert('This function is not implemented yet');
        //this.currentDocumentEndPoint = "https://graph.microsoft.com/beta/me/drive/recent";
        //this.reloadContent();
    }

    public sites = () => {
        this.currentDocumentEndPoint = "https://graph.microsoft.com/beta/me/insights/trending?$filter=ResourceReference/type neq 'microsoft.graph.driveItem'"
        this.reloadContent();
    }

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable('theme'));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext(context => {
                
                this.setState({
                    entityId: context.entityId,
                    displayName:'',
                    location:'',
                    jobTitle: '',
                    alerts: ''
                });
            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }
        this.reloadContent();
    }

    /** 
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <TeamsComponentContext
                fontSize={this.state.fontSize}
                theme={this.state.theme}
            >

                <ConnectedComponent render={(props) => {
                    const { context } = props;
                    const { rem, font } = context;
                    const { sizes, weights } = font;
                    const styles = {
                        header: { ...sizes.title, ...weights.semibold },
                        section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
                        footer: { ...sizes.xsmall }
                    }

                    return (
                        <Surface>
                            <Panel>
                                <PanelHeader>
                                    <div style={styles.header}>Trending Documents</div>
                                </PanelHeader>
                                <PanelBody>
                                    <div style={styles.section}>
                                        <PrimaryButton onClick={this.trending}>Trending Documents</PrimaryButton>
                                        <PrimaryButton onClick={this.recent}>Recent Documents</PrimaryButton>
                                        <PrimaryButton onClick={this.sites}>Trending Sites</PrimaryButton>
                                    </div>
                                    <div className="ms-Grid">
                                        <div className="ms-Grid-row">
                                            <DocumentsComponent documentsData={this.state.documents} />
                                        </div>
                                    </div>
                                    <div style={styles.section}>
                                        <PrimaryButton onClick={this.reloadContent}>Reload Documents</PrimaryButton>
                                    </div>
                                </PanelBody>
                                <PanelFooter>
                                    <div style={styles.footer}>
                                        (C) Copyright Paul Summers
                                    </div>
                                </PanelFooter>
                            </Panel>
                        </Surface>
                    );
                }}>
                </ConnectedComponent>
            </TeamsComponentContext >
        );
    }
}