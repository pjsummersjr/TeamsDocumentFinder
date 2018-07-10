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
    constructor(props: IDocumentFinderTabProps, state: IDocumentFinderTabState) {
        super(props, state);
    }

    public getProfile():void {
        console.log('Running update profile');
        let graphEndpoint = "https://graph.microsoft.com/v1.0/me";
        var req = new XMLHttpRequest();
        req.open("GET", graphEndpoint, false);
        req.setRequestHeader("Authorization", "Bearer " + this.token);
        req.setRequestHeader("Accept", "application/json;odata.metadata=minimal;");
        req.send();

        var result = JSON.parse(req.responseText);
        this.setState({
            displayName: result.displayName,
            location: result.officeLocation,
            jobTitle: result.jobTitle,
            alerts: 'Refreshed profile information'
        });
    }

    private tokenIsValid = () => {
        return (this.token && this.token.length > 0);
    }
    public loadDocuments = () => {
        console.debug('Calling loadDocuments');
        if(!this.tokenIsValid()) {
            console.debug('No token. Refreshing...');
            this.refreshToken(this.loadDocuments, this.authError);
        }
        else {
            let graphEndpoint = "https://graph.microsoft.com/beta/me/insights/trending";
            var req = new XMLHttpRequest();
            req.open("GET", graphEndpoint, true);
            req.setRequestHeader("Authorization", "Bearer " + this.token);
            req.setRequestHeader("Accept", "application/json;odata.metadata=minimal;");

            req.onload = () => {
                this.setState({
                    documents: req.responseText
                })
            }

            req.send();
        }
    }

    public loadContent2 = () => {
        const client: GraphClient = new GraphClient();
        client.graphRequest("https://graph.microsoft.com/beta/me/insights/trending", 
            (response) => { this.setState({
                documents: response
            })},
            (error) => {this.authError(error);}
        )
    }

    public updateProfile = () => {
        console.debug('Calling updateProfile');
        if(!this.tokenIsValid()) {
            console.debug('No token. Refreshing...');
            this.refreshToken(this.updateProfile, this.authError)
        }
        else {
            console.debug('Have the access token. Calling the Graph');
            let graphEndpoint = "https://graph.microsoft.com/v1.0/me";
         
            var req = new XMLHttpRequest();
            req.open("GET", graphEndpoint, true);
            req.setRequestHeader("Authorization", "Bearer " + this.token);
            req.setRequestHeader("Accept", "application/json;odata.metadata=minimal;");
            
            req.onload = () => {
                let result = JSON.parse(req.responseText);

                this.setState({
                    displayName: result.displayName,
                    location: result.officeLocation,
                    jobTitle: result.jobTitle,
                    alerts: 'Profile information successfully retrieved'
                });
            }
            req.send();
        }
    }

    public authError = (error) => {
        this.setState({
            displayName: 'Not found',
            location: 'Not found',
            jobTitle: 'Not found',
            alerts: `Could not retrieve profile information: ${error}`
        });
    }

    public refreshToken = (successMethod, errorMethod) => {
        console.log('Fetching token');
        let ls = window.localStorage;
        if(ls){
            console.debug('Local storage is active'); 
            if (ls.getItem('authtoken')) {
                console.debug('Token is cached. Returning');
                this.token = ls.getItem('authtoken');
                successMethod();
            }
            else {
                console.debug('Token is invalid or not cached. Retrieving new token.');
                this.handleAuth(
                    (token) => {
                        this.token = token;
                        ls.setItem('authtoken', token);
                        successMethod();
                    }, errorMethod);
            }
        } else {
            this.handleAuth((token) => {
                this.token = token;
                successMethod()
            }, errorMethod);
        }
    }
    /**
     * Need to put this in a separate auth library
     */
    public handleAuth = (successMethod, errorMethod) => {
        console.debug('Getting new access token');
        microsoftTeams.authentication.authenticate({
            url: "/auth.html",
            width: 700,
            height: 500,
            successCallback: (data) => {    
                console.debug('New token retrieved. Calling back.');
                successMethod(data);       
            },
            failureCallback: function (err) {
                console.debug('Retrieval of auth token failed.');
                errorMethod(err);                
            }
        });
    }

    public trending = () => {
        this.loadDocuments();
    }

    public recent = () => {
        this.loadDocuments();
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
        this.loadDocuments();
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
                                        <PrimaryButton onClick={this.trending}>Trending</PrimaryButton>
                                        <PrimaryButton onClick={this.recent}>Recent</PrimaryButton>
                                    </div>
                                    <div className="ms-Grid">
                                        <div className="ms-Grid-row">
                                            <DocumentsComponent documentsData={this.state.documents} />
                                        </div>
                                    </div>
                                    <div style={styles.section}>
                                        <PrimaryButton onClick={this.loadContent2}>Reload Documents</PrimaryButton>
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