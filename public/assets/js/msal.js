 //MSAL FUNCTIONS
async function getToken() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
        const request = {
            account: accounts[0],
            scopes: ["email", "openid","User.ReadBasic.All", "profile", "user.read", "Sites.Read.All", "Sites.ReadWrite.All"] // Replace with necessary scopes
        };
        try {
            const response = await msalInstance.acquireTokenSilent(request);
            return response.accessToken;
        } catch (error) {
            console.error("Token acquisition error:", error);
            if (error instanceof msal.InteractionRequiredAuthError) {
                // Fallback to interactive token acquisition if silent acquisition fails
                const loginResponse = await msalInstance.loginRedirect(request);
                return loginResponse.accessToken;
            } else {
                showModal("Error", "Failed to acquire user token. Please refresh the page and try again.");
                throw error;
            }

        }
    } else {
        throw new Error("No accounts found.");
    }
}
async function checkLoginStatus() {
    const redirectResponse = await msalInstance.handleRedirectPromise();

    if (redirectResponse) {
        console.log("Login successful:", redirectResponse);
        getAccountInfo();
    }
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
        accountName = accounts[0].name;
        email = accounts[0].username;
        var name = accountName.split(",");
        var firstName = (name[name.length - 1]).trim();

        $('.username').text(firstName);
        getAccountInfo();
        $('#landing').addClass('d-none');
        $('#starting').removeClass('d-none');
        return true;
    } else {
        $('#landing').removeClass('d-none');
        $('#starting').addClass('d-none');
        return false;
    }
    
        
}
function setupLogger(){
    // Enable logging
    msalInstance.setLogger({
        loggerCallback: (level, message, containsPii) => {
            console.log(message);
        },
        piiLoggingEnabled: false, // Avoid logging sensitive info
        logLevel: "verbose" // Options: error, warning, info, verbose
    });

}
