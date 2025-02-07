import { PublicClientApplication } from "@azure/msal-browser"


const aadClientId = process.env.ENTRA_APPID ?? "";
if (aadClientId === undefined || aadClientId === "") throw new Error("Test app Entra client ID is missing. Please ensure it's defined in the .env file.");
var redirectUri = window.location.pathname;
const msalParams = {
    auth: {
        authority: "https://login.microsoftonline.com/common",
        clientId: aadClientId,
        redirectUri: redirectUri
    },
}


/**
 * Combines an arbitrary set of paths ensuring and normalizes the slashes
 *
 * @param paths 0 to n path parts to combine
 */
export function combine(...paths : string[]) {
    return paths
        .map(path => path.replace(/^[\\|/]/, "").replace(/[\\|/]$/, ""))
        .join("/")
        .replace(/\\/g, "/");
}

let username = "";
let donotrefresh = false;

const app = new PublicClientApplication(msalParams);
app.initialize().then(async ()=>
{
    const res = await app.handleRedirectPromise()
    console.log("handleRedirectPromise completed")
    return res;
}).then(handleResponse)
.catch((error) => {
    console.error(error);
    donotrefresh = true;
});

function selectAccount () {

    /**
     * See here for more info on account retrieval: 
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */

    const currentAccounts = app.getAllAccounts();

    if (currentAccounts.length === 0) {
        return;
    } else if (currentAccounts.length > 1) {
        // Add your account choosing logic here
        console.warn("Multiple accounts detected.");
    } else if (currentAccounts.length === 1) {
        username = currentAccounts[0].username;
    }
}

function handleResponse(response : any) {
    if (response !== null) {
        username = response.account.username;
    } else {
        selectAccount();
    }
}

export async function getToken(command : { resource: string, type: string }) {

    let authParams : any = null;

    switch (command.type) {
        case "Default":
            authParams = { scopes: ["User.Read"] };
            break;
        case "SharePoint":
        case "SharePoint_SelfIssued":
            authParams = { scopes: [`${combine(command.resource, ".default")}`] };
            break;
        default:
            break;
    }

    return getTokenImpl(authParams);
}

export async function getTokenImpl(authParams:any)
{
    let accessToken = "";
    if(!authParams)
       return "";

    if(donotrefresh)
    {
        console.log('do not refresh');
        return "";
    }

    try {
        authParams.account = app.getAccountByUsername(username);
        // see if we have already the idtoken saved
        const resp = await app.acquireTokenSilent(authParams);
        accessToken = resp.accessToken;
    } catch (e) {
        // per examples we fall back to popup
        console.log("falling back to popup for token redirection");
        app.acquireTokenRedirect(authParams)
    }

    return accessToken;
}