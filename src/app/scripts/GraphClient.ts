
export default class GraphClient {
    private token: any = "";
    private currentUrl: string = "";

    private tokenIsValid = () => {
        return (this.token && this.token.length > 0);
    }
    public refreshToken = (successMethod, errorMethod, successCallback) => {
        console.log('Fetching token');
        let ls = window.localStorage;
        if(ls){
            console.debug('Local storage is active'); 
            if (ls.getItem('authtoken')) {
                console.debug('Token is cached. Returning');
                this.token = ls.getItem('authtoken');
                successMethod(this.currentUrl, successCallback, errorMethod);
            }
            else {
                console.debug('Token is invalid or not cached. Retrieving new token.');
                this.handleAuth(
                    (token) => {
                        this.token = token;
                        ls.setItem('authtoken', token);
                        successMethod(this.currentUrl, successCallback, errorMethod);
                    }, errorMethod);
            }
        } else {
            this.handleAuth((token) => {
                this.token = token;
                successMethod(this.currentUrl, successCallback, errorMethod);
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
        console.debug('Making a GraphClient.graphRequest');
        this.currentUrl = url;
        if(!this.tokenIsValid()) {
            console.debug('No token. Refreshing.');
            this.refreshToken(this.graphRequest, fail, success);
        }
        else {
            var req = new XMLHttpRequest();
            req.open("GET", this.currentUrl, true);
            req.setRequestHeader("Authorization", "Bearer " + this.token);
            req.setRequestHeader("Accept", "application/json;odata.metadata=minimal;");

            req.onload = () => {
                success(req.responseText);
            }

            req.send();
        }
    }

/*     public batchRequest = (requests: string[], success, fail) => {
        console.debug('batchRequest');
        this.currentUrl = "https://graph.microsoft.com/v1.0/$batch";
        if(!this.tokenIsValid()) {
            console.debug('No token. Refreshing.');
            this.refreshToken(this.batchRequest, fail, success);
        }
        else {
            var req = new XMLHttpRequest();
            req.open("POST", this.currentUrl, true);
            req.setRequestHeader('Authorization', "Bearer " + this.token);
            req.setRequestHeader("Accept", "application/json;odata.metadata=minimal;");

        }
    } */
}