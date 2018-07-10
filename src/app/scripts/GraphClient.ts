
export default class GraphClient {
    private token: any = "";

    private tokenIsValid = () => {
        return (this.token && this.token.length > 0);
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
    
    public graphRequest = (url: string, success, fail) => {
        console.debug('Calling loadDocuments');
        if(!this.tokenIsValid()) {
            console.debug('No token. Refreshing...');
            this.refreshToken(this.graphRequest, fail);
        }
        else {
            var req = new XMLHttpRequest();
            req.open("GET", url, true);
            req.setRequestHeader("Authorization", "Bearer " + this.token);
            req.setRequestHeader("Accept", "application/json;odata.metadata=minimal;");

            req.onload = () => {
                success(req.responseText);
            }

            req.send();
        }
    }
}