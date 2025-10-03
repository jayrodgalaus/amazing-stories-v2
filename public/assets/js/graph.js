//GRAPH FUNCTIONS
async function saveToSharePoint(listName, data, attachments = []) {
    // URL to create the list item
    const url = `https://graph.microsoft.com/v1.0/sites/dxcportal.sharepoint.com:/sites/ITOEECoreTeam:/lists/${listName}/items`;
    const token = await getToken();

    if (token) {
        try {
            // Step 1: Save the main list item
            const response = await fetch(url, {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ fields: data })
            });

            if (!response.ok) {
                throw new Error(`Error saving to SharePoint: ${response.statusText}`);
            }

            const result = await response.json();
            console.log(result)
            console.log("List item saved successfully");

            // Step 2: Upload attachments (if any)
            if (attachments.length > 0) {
                const itemId = result.id; // Get the ID of the created list item
                const attachmentBaseUrl = `https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/web/lists/GetByTitle('${listName}')/items(${itemId})/AttachmentFiles/add(FileName='`;
                const sharePointToken = await getSharePointToken(); // Get a new token for SharePoint API
                for (const attachment of attachments) {
                    const attachmentName = attachment.name;
                    const fileData = await attachment.arrayBuffer(); // Read the file as binary data

                    const uploadResponse = await fetch(`${attachmentBaseUrl}${attachmentName}')`, {
                        method: "POST",
                        headers: {
                            "Authorization": `Bearer ${sharePointToken}`,
                            "Accept": "application/json;odata=verbose"
                        },
                        body: fileData
                    });

                    if (!uploadResponse.ok) {
                        throw new Error(`Error uploading attachment: ${uploadResponse.statusText}`);
                    }

                    console.log(`Attachment '${attachmentName}' uploaded successfully`);
                }
            }

            // Return the result of the saved list item
            return result;

        } catch (error) {
            console.error(error);
            throw new Error(`Failed to save item or attachments: ${error.message}`);
        }
    } else {
        throw new Error("Failed to retrieve authorization token");
    }
}
async function saveToSharePointBatch(listName, dataArray) {
    const url = "https://graph.microsoft.com/v1.0/$batch"; // Graph API batch endpoint
    const token = await getToken();

    if (token) {
        const batchSize = 20; // Maximum requests per batch
        const batches = [];
        let totalEntries = dataArray.length; // Total number of entries to insert
        let successCount = 0; // Counter for entries with 201 status

        // Split dataArray into batches of 20
        for (let i = 0; i < dataArray.length; i += batchSize) {
            const batch = dataArray.slice(i, i + batchSize);
            batches.push(batch);
        }

        for (const batch of batches) {
            const batchRequests = batch.map((item, index) => ({
                id: `${index + 1}`, // Unique ID for each request
                method: "POST",
                url: `/sites/dxcportal.sharepoint.com:/sites/ITOEECoreTeam:/lists/${listName}/items`,
                body: { fields: item },
                headers: {
                    "Content-Type": "application/json"
                }
            }));

            // Prepare batch payload
            const batchPayload = { requests: batchRequests };

            // Send batch request
            const response = await fetch(url, {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(batchPayload)
            });

            if (!response.ok) {
                console.error(`Error in batch request: ${response.statusText}`);
            } else {
                const result = await response.json();

                // Check individual responses for errors and count successes
                result.responses.forEach((res) => {
                    if (res.status === 201) {
                        successCount++;
                    } else if (res.status === 400) {
                        console.error(`Error in item ${res.id}: ${res.body.error.message}`);
                    }
                });
            }
        }

        // Compare counters and trigger appropriate response
        if (successCount === totalEntries) {
            showModal("Success", "All entries were saved successfully!","success");
        } else if (successCount > 0) {
            showModal(
                "Partial Success",
                `${successCount} out of ${totalEntries} entries were saved successfully. Check logs for errors.`,
                'fail'
            );
        } else {
            showModal("Error", "No entries were saved. Please check the logs for details.",'fail');
        }
    }
}
async function getFromSharePoint(listName, fields, conditions) {
    const url = `https://graph.microsoft.com/v1.0/sites/dxcportal.sharepoint.com:/sites/ITOEECoreTeam:/lists/${listName}/items`;
    const token = await getToken();

    if (token) {
        try {
            // Build the query string for filtering conditions
            let filterQuery = [];

            // Add conditions filters
            if (conditions && conditions.length > 0) {
                conditions.forEach(cond => {
                    filterQuery.push(`fields/${cond.field} eq '${cond.value}'`);
                });
            }

            // Construct the URL with $expand and $filter
            let requestUrl = `${url}?$expand=fields`;
            if (filterQuery.length > 0) {
                requestUrl += `&$filter=${filterQuery.join(" and ")}`;
            }
            console.log(requestUrl);
            // Fetch items from the SharePoint list
            const response = await fetch(requestUrl, {
                method: "GET",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                }
            });

            if (!response.ok) {
                throw new Error(`Error fetching from SharePoint: ${response.statusText}`);
            }

            const result = await response.json();
            var filteredItems = [];
            
            
            // Post-process to return only the specified fields, including 'createdBy'
            if (!fields || fields.length === 0) {
                filteredItems = result.value.map(item => ({
                    ...item.fields, // Spread operator to keep all fields
                    createdBy: item.createdBy.user // Include createdBy details
                }));
            } else {
                filteredItems = result.value.map(item => {
                    const processedItem = {};
                    fields.forEach(field => {
                        processedItem[field] = item.fields[field]; // Extract only the specified fields
                    });
                    // Add 'createdBy' information
                    processedItem.createdBy = item.createdBy.user;
                    return processedItem;
                });
            }

            console.log("Filtered Items:", filteredItems);
            return filteredItems; // Return the processed items
        } catch (error) {
            console.error("Error retrieving data from SharePoint:", error);
            throw error;
        }
    } else {
        throw new Error("Failed to obtain authentication token.");
    }
}
    /* 
    listName = name of list in sharepoint
    data = array of objects to update. each object must have an "id" field
    */
async function getListWithSP_API(listName, fields=[], conditions=[], author = null) {
    const token = await getSharePointToken();
    if (!token) {
        throw new Error("Failed to obtain authentication token.");
        showModal("Error","Failed to obtain authentication token. Please refresh the page and try again.")
    };

    try {
        const url = `https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/web/lists/GetByTitle('${listName}')/items`;
        let filterQuery = [];

        // Filter by Created By (AuthorId)
        if (author) {
            
            const userID = await getUserDetailsFromEmail(author ? author : email);
            filterQuery.push(`AuthorId eq ${userID}`);
        }

        // Add other conditions (direct properties, NOT fields/)
        if (conditions && conditions.length > 0) {
            conditions.forEach(cond => {
                if(cond.field == 'SUBSL'){
                    filterQuery.push(`${cond.field} eq '${encodeURIComponent(cond.value)}'`);
                }else
                    filterQuery.push(`${cond.field} eq '${cond.value}'`);
            });
        }
        if(listName == splist)
            filterQuery.push("Is_x0020_Deleted eq false"); // Ensure we only get non-deleted items

        // Construct final request URL
        let requestUrl = url;
        if (filterQuery.length > 0) {
            requestUrl += `?$filter=${filterQuery.join(" and ")}`;
        }
        const response = await fetch(requestUrl, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": "application/json",
                "Accept": "application/json;odata=verbose"
            }
        });

        if (!response.ok) {
            throw new Error(`Error fetching from SharePoint: ${response.statusText}`);
        }

        const result = await response.json();
        const items = result.d?.results || []; // Ensure we reference d.results

        // Retrieve CreatedBy (AuthorId) details separately
        const processedItems = await Promise.all(items.map(async item => {
            return {
                ...item, // Keep all direct properties
            };
        }));

        // console.log("Filtered Items:", processedItems);
        return processedItems; 
    } catch (error) {
        console.error("Error retrieving data from SharePoint:", error);
        showModal("Error","Error retrieving data from SharePoint")
        throw error;
    }
}
async function updateSPItem(listName, id, fieldsArray) {
    const token = await getSharePointToken();
    if (!token) {
        showModal("Error", "Failed to obtain authentication token. Please refresh the page and try again.");
        throw new Error("Failed to obtain authentication token.");
    }

    try {
        const url = `https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/web/lists/GetByTitle('${listName}')/items(${id})`;

        // Convert the array to an object for SharePoint API
        const fieldsObject = fieldsArray.reduce((acc, field) => {
            acc[field.name] = field.value;
            return acc;
        }, {});

        const response = await fetch(url, {
            method: "PATCH",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": "application/json",
                "X-HTTP-Method": "MERGE",
                "IF-MATCH": "*",
                "Accept": "application/json;odata=verbose"
            },
            body: JSON.stringify(fieldsObject)
        });

        if (!response.ok) {
            throw new Error(`Error updating SharePoint item: ${response.statusText}`);
        }

        console.log("Item updated successfully");
    } catch (error) {
        console.error("Error updating item:", error);
        showModal("Error", "Error updating item");
        throw error;
    }
}


async function updateToSharePoint(listName, data) {
    const url = `https://graph.microsoft.com/v1.0/sites/dxcportal.sharepoint.com:/sites/ITOEECoreTeam:/lists/${listName}/items`;
    const token = await getToken();

    if (token) {
        try {
            // Validate that data contains an 'id' column
            if (!data.every(item => item.id)) {
                throw new Error("Each item in 'data' must have an 'id' field.");
            }

            // Prepare batch request payload
            const batchRequest = {
                requests: data.map(item => {
                    const requestUrl = `${url}/${item.id}`; // Use 'id' field for identifying the item
                    return {
                        id: item.id, // Unique identifier for batch request
                        method: "PATCH",
                        url: requestUrl,
                        headers: {
                            "Authorization": `Bearer ${token}`,
                            "Content-Type": "application/json"
                        },
                        body: {
                            fields: { ...item } // Spread the key-value pairs from the current item
                        }
                    };
                })
            };

            // Send batch request
            const batchResponse = await fetch(`https://graph.microsoft.com/v1.0/$batch`, {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(batchRequest)
            });

            if (!batchResponse.ok) {
                throw new Error(`Error in batch update: ${batchResponse.statusText}`);
            }

            const result = await batchResponse.json();
            console.log("Batch update response:", result);

            console.log("All items updated successfully.");
            return result;
        } catch (error) {
            console.error("Error updating data to SharePoint:", error);
            throw error;
        }
    } else {
        throw new Error("Failed to obtain authentication token.");
    }
}
async function signIn() {
    try {
        // Handle the redirect after login
        const redirectResponse = await msalInstance.handleRedirectPromise();

        if (redirectResponse) {
            console.log("Login successful:");
            checkLoginStatus(); // Ensure this runs after a successful login
            return; // Stop further execution after reload
        }

        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) {
            // Initiate login redirect
            await msalInstance.loginRedirect({
                scopes: ["email", "openid","User.ReadBasic.All", "profile", "user.read", "Sites.Read.All", "Sites.ReadWrite.All"] // Initial scopes
                
            });
            checkLoginStatus(); // Ensure this runs after a successful login
            window.location.reload();
        } else {
            
            checkLoginStatus(); // Account already exists
        }
    } catch (error) {
        console.error("Login error:", error);
        showModal("Login Error", "There was an error signing in. Please refresh the browser and try again.");
    }
}
async function getAccountInfo() {
    const token = await getToken();
    if (token) {
        const accounts = msalInstance.getAllAccounts();
        accountName = (accounts[0].name).trim();
        email = (accounts[0].username).trim();

        console.log("Account Name:", accountName, "Email:", email);
    }
}

// Function to fetch users from the tenant directory
async function fetchTenantUsers(searchTerm) {
    try {
        // Acquire the access token using MSAL
        const account = msalInstance.getActiveAccount(); // Ensure there's a signed-in user
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["User.ReadBasic.All"],
            account: account,
        });

        const accessToken = tokenResponse.accessToken;

        // Call Microsoft Graph API to fetch users with filtering on first and last name
        const graphEndpoint = `https://graph.microsoft.com/v1.0/users?$select=id,displayName,givenName,surname,mail&$filter=startswith(givenName,'${searchTerm}') or startswith(surname,'${searchTerm}') or startswith(displayName,'${searchTerm}')`;

        const response = await fetch(graphEndpoint, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                "Content-Type": "application/json",
            },
        });

        if (!response.ok) {
            throw new Error(`Error fetching data: ${response.statusText}`);
        }

        const data = await response.json();

        // Parse and return the relevant fields
        return data.value.map(user => ({
            id: user.id,
            name: user.displayName,
            firstName: user.givenName || "No First Name",
            lastName: user.surname || "No Last Name",
            email: user.mail || "No Email",
        }));
    } catch (error) {
        console.error("Error fetching user data: ", error);
        return [];
    }
}
