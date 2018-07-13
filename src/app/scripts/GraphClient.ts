
export default class GraphClient {

    public getToken(success, error): any {

        let ls = window.localStorage;
        if(ls && ls.getItem('authtoken')){
            console.debug('Token in local storage');
            
            if(ls.getItem('authtoken')) { 
                success(ls.getItem('authtoken'));
            } else {
                error("Error retrieving auth token from local storage even though I checked to make sure it was already there.");
            }
        }
        else {
            return this.handleAuth(
                (response) => {
                    ls.setItem('authtoken', response);
                    success(response);
                },
                (error) => {
                    error(`Error retrieving getting new auth token.\n${error}`);
                }
            );
        }
    }


    /**
     * Need to put this in a separate auth library
     */
    public handleAuth(successMethod, errorMethod) {
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
    
        this.getToken(
            (token) => {

                var req = new XMLHttpRequest();
                req.open("GET", url, true);
                req.setRequestHeader("Authorization", "Bearer " + token);
                req.setRequestHeader("Accept", "application/json;odata.metadata=minimal;");
        
                req.onload = () => {
                    success(req.responseText);
                }
        
                req.send();

            },
            (e) => {
                console.error(`Error retrieving the documents.\n${e}`);
            }

        );
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