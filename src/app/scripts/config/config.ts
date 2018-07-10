export default class AppConfig {

    public static something: string = "";
   public static permissionScopes: string[] = [
            "https://graph.microsoft.com/user.read", 
            "https://graph.microsoft.com/group.read.all", 
            "https://graph.microsoft.com/Sites.Read.All",
            "https://graph.microsoft.com/Sites.ReadWrite.All"
        ];
    
    public static graphApiKey: string = "d9521502-2ab8-4b93-9b80-3e8aafe69b69";
    public static graphAuthEndPoint: string = "https://login.microsoftonline.com/common";
}