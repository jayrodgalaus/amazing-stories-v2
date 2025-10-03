//SHAREPOINT FUNCTIONS
async function getSharePointToken(){
    try{
        const accounts = msalInstance.getAllAccounts();
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["https://dxcportal.sharepoint.com/.default"], // SharePoint API scope
            account: accounts[0], // Use the first account
        });
        return tokenResponse.accessToken;
    } catch (error) {
        console.error("Error acquiring token:", error);
        if (error instanceof msal.InteractionRequiredAuthError) {
            // Fallback to interactive token acquisition if silent acquisition fails
            const loginResponse = await msalInstance.loginRedirect({
                scopes: ["https://dxcportal.sharepoint.com/.default"], // SharePoint API scope
            });
            return loginResponse.accessToken;
        } else {
            showModal("Notice", "Session expired. Please refresh the page and try again.");
            throw error;
        }
    }
    
}
async function getSiteAndDriveDetails() {
    console.log("Fetching site and drive details.")
    const account = msalInstance.getAllAccounts()[0];
    tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: ["Sites.Read.All", "Sites.ReadWrite.All"],
        account: account
    });
    // ACTION REQUIRED
    // Get the Site ID - change to new site
    const siteResponse = await fetch("https://graph.microsoft.com/v1.0/sites?search=ITOEECoreTeam", {
        headers: {
            Authorization: `Bearer ${tokenResponse.accessToken}`
        }
    });

    const siteData = await siteResponse.json();
    siteId = siteData.value[0].id;

    // Get the Drive ID for the document library
    const driveResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
        headers: {
            Authorization: `Bearer ${tokenResponse.accessToken}`
        }
    });

    const driveData = await driveResponse.json();
    driveId = driveData.value[0].id; // Assuming the first drive is the document library

    // return { siteId, driveId, tokenResponse };
}
async function getGraphAttachments(siteId, listId, itemId, accessToken) {
    const endpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/attachments`;

    try {
        const response = await fetch(endpoint, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Content-Type": "application/json"
            }
        });

        if (!response.ok) {
            throw new Error(`Error fetching attachments: ${response.statusText}`);
        }

        const result = await response.json();
        console.log("Attachments:", result.value);
        return result.value; // Contains attachment details including download URLs
    } catch (error) {
        console.error("Error retrieving attachments:", error);
    }
}
async function getUserDetailsById(id) {
    try{
        const token = await getSharePointToken();
        const userResponse = await fetch(`https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/web/siteusers?$filter=Id eq ${id}&$select=Id,Title,Email`, {
            headers: { "Authorization": `Bearer ${token}` , "Accept": "application/json;odata=verbose"}
        });
        const userData = await userResponse.json();
        return userData.d.results[0]; // Assuming only one match
    }catch(error){
        console.error("Failed to fetch user details:", error);
        return null;
    }
    
}
async function getUserDetailsFromEmail(email) {
    const token = await getSharePointToken();
    const userIdURL = `https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/web/siteusers?$filter=Email eq '${email}'&$select=Id,Title,Email`;
    const userResponse = await fetch(userIdURL, {
                method: "GET",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json",
                    "Accept": "application/json;odata=verbose"
                }
            });
    const userResult = await userResponse.json();
    return userResult.d.results[0].Id;
}
async function userLevel(){
    const token = await getSharePointToken();
    if (!token) {
        showModal("Error","Failed to obtain authentication token. Please refresh the page and try again.");
        throw new Error("Failed to obtain authentication token.");
    };

    try {
        /* 1 = Super, 2 = SPOC, 3 = others */
        let defaultaccess = {type: 3, subsl: "Others"};
        const url = `https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/web/lists/GetByTitle('Amazing Stories POC List')/items?$filter=POCName/EMail eq '${email}'&$select=SUBSL/Title,POCName/Title,POCName/EMail,POCName/Id&$expand=SUBSL,POCName`;
        
        const response = await fetch(url, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": "application/json",
                "Accept": "application/json;odata=verbose"
            }
        });

        if (!response.ok) {
            console.log(`No access to POC List`);
            $('.spoc-element').remove();
            $('.super-element').remove();
            $('.other-element').removeClass("d-none");
            // throw new Error(`Error fetching from SharePoint: ${response.statusText}`);
        }else{
            const result = await response.json();
            const items = result.d?.results || []; // Ensure we reference d.results

            // Retrieve CreatedBy (AuthorId) details separately
            const processedItems = await Promise.all(items.map(async item => {
                return {
                    ...item, // Keep all direct properties
                };
            }));
            
            if(processedItems.length > 0){
                let user = processedItems[0];
                defaultaccess.subsl = user.SUBSL.Title;
                if(defaultaccess.subsl == 'GIS' || defaultaccess.subsl == 'C&I'){
                    defaultaccess.type = 1; // Super Admin
                    $('.super-element').removeClass("d-none");
                }else {
                    defaultaccess.type = 2; // SPOC
                    $('.super-element').remove();
                    //remove subsl options
                    $("#subslDropdown option, #updateSubslDropdown option").not("[value='" + defaultaccess.subsl + "']").remove();
                }
                $('.spoc-element').removeClass("d-none");
            }else{
                //remove everything
                $('.spoc-element').remove();
                $('.super-element').remove();
                $('.other-element').removeClass("d-none");
            }
        }
        
        return defaultaccess; 
        

    } catch (error) {
        console.error("Error retrieving data from SharePoint:", error);
        showModal("Error","Error retrieving data from SharePoint.")
        throw error;
    }
}
async function getFormDigest() {
    try {
        const token = await getSharePointToken();
        const response = await fetch(`https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/contextinfo`, {
            method: "POST",
            headers: {
                "Accept": "application/json;odata=verbose",
                "Authorization": "Bearer "+token
            }
        });
        const data = await response.json();
        return data.d.GetContextWebInformation.FormDigestValue;
    } catch (error) {
        console.error("Error fetching request digest: ", error);
        return null;
    }
}
async function addAttachments(itemId, files) {
    const siteUrl = "";  // Replace with your site URL
    const listName = "YourListName";  // Replace with your list name

    try {
        const formDigest = await getFormDigest();  // Ensure you have a valid form digest

        for (const file of files) {
            const fileName = file.name;
            const token = await getSharePointToken();
            const response = await fetch(`https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/web/lists/getbytitle('Amazing Stories entries')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`, {
                method: "POST",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": formDigest,
                    "Authorization": "Bearer "+token,
                },
                body: file
            });

            if (!response.ok) {
                throw new Error(`Upload failed for ${fileName}: ${response.statusText}`);
                showModal("Error", `Upload failed for ${fileName}: ${response.statusText}`);
            }

            console.log(`Attachment uploaded successfully: ${fileName}`);
        }
    } catch (error) {
        console.error("Error uploading attachments:", error);
        showModal("Error", "Error uploading attachments: " + error.message);
    }
}

async function deleteAttachments(listName, itemId, fileNames) {
    const formDigest = await getFormDigest();
    if (!formDigest) {
        console.error("Could not retrieve request digest.");
        return;
    }
    const token = await getSharePointToken();
    for (let fileName of fileNames) {
        let endpoint = `https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/web/lists/GetByTitle('${listName}')/items(${itemId})/AttachmentFiles('${fileName}')`;

        try {
            const response = await fetch(endpoint, {
                method: "DELETE",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": formDigest,
                    "IF-MATCH": "*",
                    "X-HTTP-Method": "DELETE",
                    "Authorization": "Bearer "+token
                }
            });

            if (!response.ok) {
                throw new Error(`Failed to delete ${fileName}: ${response.statusText}`);
            }

            console.log(`Deleted: ${fileName}`);
        } catch (error) {
            console.error(`Error deleting attachment (${fileName}):`, error);
        }
    }
}
