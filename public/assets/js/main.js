const msalConfig = {
    auth: {
        clientId: "2211ce0f-1187-409a-9c1d-baf8a0dd577f", // Replace with your client ID
        authority: "https://login.microsoftonline.com/93f33571-550f-43cf-b09f-cd331338d086", // Replace with your tenant ID
        redirectUri: "https://amazing-stories-dev.vercel.app/" // Replace with your redirect URI
    }

};
const msalInstance = new msal.PublicClientApplication(msalConfig);
var siteId, driveId, tokenResponse;
var authorId, accountName, email, access;
const splist = "Amazing Stories entries dev";
var spItems = [];
const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
var selectedEntries = [];
var deferredPrompt;

async function init(){
    setCurrentMonth();
    const observer = new IntersectionObserver((entries, observer) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const el = entry.target;
                const bgUrl = el.getAttribute('data-bg');
                el.style.backgroundImage = `url('${bgUrl}')`;
                observer.unobserve(el); // Lazy load only once
            }
        });
    });

    document.querySelectorAll('.lazy-bg').forEach(el => {
        observer.observe(el);
    });
    document.querySelectorAll('.lazy-bg.highres').forEach(el => {
        const img = new Image();
        const url = el.getAttribute('data-bg');
        img.onload = () => {
            requestAnimationFrame(() => {
                el.style.backgroundImage = `url(${url})`;
                el.classList.add('loaded');
            });
        };
        img.src = url;
    });
    let status = await checkLoginStatus();
    if(status){
        getSiteAndDriveDetails();
        addFormValidation();
        // embedPowerBIReport();
        registerServiceWorker();
        checkIfInstalled();
        //check access level
        authorId = await getUserDetailsFromEmail(email);
        access = await userLevel();
        dummydata();
        
    }

    $('#home-page').css('opacity', '1');
}
//PWA Functions
function registerServiceWorker() {    
    if ('serviceWorker' in navigator) {
        window.addEventListener('load', () => {
            navigator.serviceWorker.register('/sw.js')
            .then(reg => console.log('Service Worker registered:'))
            .catch(err => console.error('Service Worker registration failed:', err));
        });
    }
}

function checkIfInstalled(){
    const isStandalone = window.matchMedia('(display-mode: standalone)').matches ||
    window.navigator.standalone === true; // for iOS
    if (isStandalone) {
        $('#installBtn').hide(); // or .remove()
        $(window).on('contextmenu', function (e) {
            e.preventDefault();
            console.log("Right-click disabled (jQuery)");
        });
        console.log("This is the standalone app");
    }
}

//ADD ENTRY FUNCTIONS
function setCurrentMonth() {
    // Get the current month index (0 for January, 1 for February, etc.)
    const currentMonthIndex = new Date().getMonth();

    // Get the monthDropdown element
    const monthDropdown = document.getElementById("monthDropdown");

    // Set the selected option to the current month
    monthDropdown.selectedIndex = currentMonthIndex;

    //disable future months
    $("#monthDropdown option").each(function () {
        const optionText = $(this).val(); // or use .text()
        const optionIndex = monthNames.indexOf(optionText);

        if (optionIndex > currentMonthIndex) {
            $(this).prop("disabled", true);
        }
    });
    

}
    // Function to update fileInput with the remaining files
function updateFileInput(filesArray, id) {
    const fileInput = document.getElementById(id);
    const dataTransfer = new DataTransfer(); // Create a new DataTransfer object
    
    filesArray.forEach(file => {
        dataTransfer.items.add(file); // Add the remaining files to the DataTransfer object
    });
    
    fileInput.files = dataTransfer.files; // Update fileInput's FileList
}

//ENTRIES FUNCTIONS
async function viewItem(id){
    let entryItem = spItems[id];
    let title = entryItem.Title;
    let year = entryItem.Year;
    let month = entryItem.Month;
    let subsl = entryItem.SUBSL;
    let account = entryItem.Account;
    let team = entryItem.Team;
    let recognition = entryItem.Individual ? "Individual" : "Team";
    let recipients = entryItem.Recipients.split('/ ');
    let recipientHTML = "";
    let worktype = entryItem.Worktype;
    let challenge = entryItem.Challenge.replace(/\n/g, "<br>");
    let help = entryItem.Help.replace(/\n/g, "<br>");
    let impact = entryItem.Impact.replace(/\n/g, "<br>");
    let uniqueness = entryItem.Uniqueness ? entryItem.Uniqueness.replace(/\n/g, "<br>") : "N/A";
    let author = await getUserDetailsById(entryItem.AuthorId);
    let createdBy = author ? author.Title : "Unknown";
    let createdOn = new Date(`${entryItem.Created}`);
    const options = { year: "numeric", month: "long", day: "numeric" };
    createdOn = createdOn.toLocaleDateString("en-US", options);
    let attachmentID = entryItem.Attachments ? entryItem.ID : null;
    let amplified = entryItem.Amplified;
    /* let title = entry.data('title');
    let year = entry.data('year');
    let month = entry.data('month');
    let subsl = entry.data('subsl');
    let account = entry.data('account');
    let team = entry.data('team');
    let recognition = entry.data('recognition') ? "Individual" : "Team";
    let recipients = entry.data('recipients').split('/ ');
    let recipientHTML = "";
    let worktype = entry.data('worktype');
    let challenge = entry.data('challenge');
    let help = entry.data('help');
    let impact = entry.data('impact');
    let uniqueness = entry.data('uniqueness') ? entry.data('uniqueness') : "N/A";
    let author = await getUserDetailsById(entry.data('author'));
    let createdBy = author ? author.Title : "Unknown";
    let createdOn = entry.data('createdOn');
    let attachmentID = entry.data('attachments'); */
    let logo = getAccountLogo(account);
    $('#entryInfoEntry').text(title);
    $('#entryInfoYear').text(year);
    $('#entryInfoMonth').text(month);
    $('#entryInfoSubSL').text(subsl);
    $('#entryInfoTeam').text(team);
    $('#entryInfoCreatedBy').text(createdBy);
    $('#entryInfoCreatedOn').text(createdOn);
    $('#entryInfoLogo').attr('src', logo);
    recipients.forEach(recipient => {
        recipientHTML += `${recipient}<br>`
    });
    $('#entryInfoRecipients').html(recipientHTML);
    $('#entryInfoLogo').attr("src", getAccountLogo(account));
    $('#entryInfoWorkType').text(worktype);
    $('#entryInfoChallenge').html(challenge);
    $('#entryInfoImpact').html(impact);
    $('#entryInfoUniqueness').html(uniqueness);
    $('#entryInfoHelp').html(help);
    if(attachmentID){
        await displayAttachments(attachmentID).then(()=>{
            $('#entryInfoImages').fadeIn();
        }).catch((error) => {
            $('#entryInfoImages').html(`<p>Error retrieving attachments</p>`)
            console.error("Error retrieving attachments:", error);
        });
    }else{
        $('#entryInfoImages').html(`<p>No attachments.</p>`)
    }
    $('#entryInfoAmplifyBtn').attr('data-amplified', amplified);
    $('#entryLoading').removeClass('d-flex').addClass('d-none');
    $('#entryInfoActions button').attr('data-id', id);
    $('#entryInfoCard').fadeIn();
}
async function getEntryById(id){
    let fields = [
        "Title","Year","Month","SUBSL","Account","Team","Individual","Recipients","Recipient_x0020_Emails","Worktype","Challenge","Help","Impact","Uniqueness","Outcome","Amplified",'Id','AuthorId', "Classification"
    ];
    let conditions = [{field:"ID", value: id}];
    let data = await getListWithSP_API(splist,fields,conditions);
    
    if(data.length > 0)
        return data[0]; // Return the first item
    else
        console.log("No data found");
    
}
async function getAllEntriesByMonth(month){
    let sort = $('#sorter').val();
    $('#entriesContainer').html(loadingHTML());
    let fields = [
        "Title","Year","Month","SUBSL","Account","Team","Individual","Recipients","Recipient_x0020_Emails","Worktype","Challenge","Help","Impact","Uniqueness","Outcome","Amplified",'Id','AuthorId', "Classification"
    ];
    let conditions = [];
    if(month != 'All'){
        conditions = [{field:"Month", value: month}];
    }
    if(access.type == 1){
        let selectedSubSL = $('#subslFilter').val();
        if(selectedSubSL != 'All'){
            conditions.push({field:"SUBSL", value: selectedSubSL});
        }
    }else if(access.type == 2){
        conditions.push({field:"SUBSL", value: access.subsl});
    }
    // conditions.push({field:"Is_x0020_Deleted", value: false});
    getListWithSP_API(splist,fields,conditions).then(data=>{
        if(data.length > 0)
            processListItems(data,month,sort); // Process the retrieved items
        else
            $('#entriesContainer').html(noContentHTML());
    }).catch((error) => {
        console.error("Error retrieving data:", error);
    });
}
async function getSelfEntriesByMonth(month){
    let sort = $('#sorter').val();
    $('#entriesContainer').html(loadingHTML());
    let fields = [
        "Title","Year","Month","SUBSL","Account","Team","Individual","Recipients","Recipient_x0020_Emails","Worktype","Challenge","Help","Impact","Uniqueness","Outcome","Amplified",'Id','AuthorId', "Classification"
    ];
    let conditions = [];
    if(month != 'All'){
        conditions = [{field:"Month", value: month}];
    }
    if(access.type == 1){
        let selectedSubSL = $('#subslFilter').val();
        if(selectedSubSL != 'All'){
            conditions.push({field:"SUBSL", value: selectedSubSL});
        }
    }else if(access.type == 2){
        conditions.push({field:"SUBSL", value: access.subsl});
    }
    
    // conditions.push({field:"Is_x0020_Deleted", value: false});
    getListWithSP_API(splist,fields,conditions,email).then(data=>{
        if(data.length > 0)
            processListItems(data,month,sort); // Process the retrieved items
        else
            $('#entriesContainer').html(noContentHTML());
    }).catch((error) => {
        console.error("Error retrieving data:", error);
    });

    
}
async function getOwnRecognition(month) {
    let sort = $('#sorter').val();
    $('#entriesContainer').html(loadingHTML());
    try {
        let fields = ["Id", "Recipient_x0020_Emails"];
        let conditions = [];
        if(month != 'All'){
            conditions = [{field:"Month", value: month}];
        }
        let data = await getListWithSP_API(splist, fields, conditions);
        if (data && data.length > 0) {
            
            let fields = [
                "Title", "Year", "Month", "SUBSL", "Account", "Team", "Individual", 
                "Recipients", "Recipient_x0020_Emails", "Worktype", "Challenge", "Help", 
                "Impact", "Uniqueness", "Outcome", "Amplified", "Id", "AuthorId", "Classification", "Submitted_x0020_By", "Created", "Attachments"
            ];
            let filteredData = data.filter(item => item.Recipient_x0020_Emails.includes(email));
            if(filteredData.length == 0){
                $('#entriesContainer').html(noContentHTML());
                return false;
            }
            let ids = filteredData.map(datum => datum.Id);
            if(month != 'All'){
                conditions = [{field:"Month", value: month}];
            }
            let filterQuery = ids.map(id => month !== 'All' ? `(ID eq ${id} and Month eq '${month}')` : `(ID eq ${id})`).join(' or ');
            let selectQuery = fields.join(',');

            const url = `https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/web/lists/getByTitle('${splist}')/items?$filter=${filterQuery}&$select=${selectQuery}`;
            const token = await getSharePointToken();
            const response = await fetch(url, {
                method: "GET",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Accept": "application/json;odata=verbose"
                }
            });

            const responseData = await response.json();
            console.log(responseData);
            if(responseData.d.results.length > 0)
                processListItems(responseData.d.results,month,sort); // Process the retrieved items
            else
                $('#entriesContainer').html(noContentHTML());
        } else {
            $('#entriesContainer').html(noContentHTML());
        }
    } catch (error) {
        console.log(error);
        showModal("Error", "Error fetching data. Please try again.");
    }
}
function getAccountLogo(account){
    let img = "";
    switch(account){
        case "ADBRI": img = "ADBRI"; break;
        case "AgGateway":  img = "AgGateway"; break;
        case "Airbus": img = "Airbus"; break;
        case "Amcor": img = "Amcor"; break;
        case "AON": img = "AON"; break;
        case "AT&T": img = "AT&T"; break;
        case "Audi": img = "Audi"; break;
        case "Avanos": img = "Avanos Medical"; break;
        case "Aviva": img = "Aviva Canada"; break;
        case "BOQ": img = "BOQ"; break;
        case "Betafence": img = "Betafence"; break;
        case "BlueScope": img = "BlueScope"; break;
        case "Calvo": img = "Calvo"; break;
        case "CIBC": img = "CIBC"; break;
        case "Coca-Cola": img = "Coca Cola"; break;
        case "DirecTV": img = "DIRECTV"; break;
        case "Duracell": img = "Duracell"; break;
        case "Downer": img = "Downer EDI Limited"; break;
        case "DXC Technology": img = "DXCTechnology"; break;
        case "Fischer": img = "Fischer"; break;
        case "Glanbia": img = "Glanbia"; break;
        case "Global Life Sciences Products and Services": img = "Global Life Sciences Products & Services"; break;
        case "Haleon": img = "Haleon"; break;
        case "Hanes Brands Inc.": img = "Hanes Brands Inc."; break;
        case "HF Sinclair": img = "HF Sinclair"; break;
        case "Hitachi": img = "Hitachi"; break;
        case "HSC BSO": img = "HSC BSO"; break;
        case "IAG": img = "IAG"; break;
        case "Japan Tobacco Inc.": img = "Japan Tobacco Inc."; break;
        case "Jollibee Foods Corporation": img = "JFC"; break;
        case "KBR": img = "KBR"; break;
        case "Kraft Heinz": img = "Kraft Heinz"; break;
        case "Leveraged": return null; break;
        case "Latitude Financial Services": img = "Latitude Financial Services"; break;
        case "Macquarie": img = "Macquarie"; break;
        case "Markem-Imaje": img = "Markem Imaje"; break;
        case "Medmix": img = "Medmix"; break;
        case "Microsoft": img = "Microsoft"; break;
        case "Nestle": img = "Nestle"; break;
        case "Nissan": img = "Nissan"; break;
        case "ONE": img = "ONE"; break;
        case "Oceana": img = "Oceana"; break;
        case "Origin": img = "Origin"; break;
        case "Philips": img = "Philips"; break;
        case "P&G": img = "P&G"; break;
        case "Pilmico": img = "Pilmico"; break;
        case "Radisson": img = "Radisson Hotel Group"; break;
        case "Ralph Lauren": img = "Ralph Lauren"; break;
        case "Sabre": img = "Sabre"; break;
        case "Serco": img = "Serco"; break;
        case "Siam Cement Group": img = "Siam Cement Group"; break;
        case "Sonneborn" : img = "Sonneborn"; break;
        case "Sotheby's": img = "Sothebys"; break;
        case "South Australian Health": img = "South Australian Health"; break;
        case "Sulzer": img = "Sulzer"; break;
        case "ThyssenKrupp": img = "ThyssenKrupp"; break;
        case "Toronto Dominion Bank": img = "Toronto Dominion Bank"; break;
        case "Tops Markets LLC": img = "Tops"; break;
        case "Transport for NSW": img = "Transport for NSW"; break;
        case "Uniper": img = "Uniper"; break;
        case "Valeo": img = "Valeo"; break;
        case "Ventia": img = "Ventia"; break;
        case "Western Sydney Airport": img = "Western Sydney Airport"; break;
        case "Western Union": img = "Western Union"; break;
        case "Westpac": img = "Westpac"; break;
        case "Whitehaven Coal": img = "Whitehaven Coal"; break;
        case "W.L. Gore": img = "W.L. Gore"; break;
        case "Worksafe Victoria": img = "WorkSafe Victoria"; break;
        case "ZF WABCO": img = "ZF Wabco"; break;
        default: return null; break;
    }
    return "assets/img/logos/"+img + ".jpg";
}
async function getAttachments(itemID) {
    try{
        const token = await getSharePointToken(); // Get a new token for SharePoint API
        const response = await fetch(`https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/Web/Lists/GetByTitle('${splist}')/items(${itemID})/AttachmentFiles`, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Accept": "application/json;odata=verbose"
            }
        });
        const data = await response.json();
        return data.d.results; // Returns an array of attachments
        
    }catch(error){
        console.log(error);
        return null;
    }
    
}
async function displayAttachments(itemID) {
    const attachments = await getAttachments(itemID);
    let html = "";

    attachments.forEach(file => {
        html += `<img src="https://dxcportal.sharepoint.com${file.ServerRelativeUrl}" data-name="${file.FileName}" alt="It seems there is a problem with this image. Please upload a replacement." style="max-width: 200px; margin: 5px;">`;
    });

    document.getElementById("entryInfoImages").innerHTML = html;
}
async function displayExistingAttachmentsInUpdate(itemID) {
    const attachments = await getAttachments(itemID);
    let previewContainer = document.getElementById("existingPreviewContainer");
    if(attachments.length == 0){
        previewContainer.innerHTML = `<p>No images.</p>`;
        return false;
    }
    attachments.forEach((file, index) => {
        // html += `<img src="https://dxcportal.sharepoint.com${file.ServerRelativeUrl}" data-name="${file.FileName}" alt="Attachment" style="max-width: 200px; margin: 5px;">`;
        
        const div = document.createElement('div'); // Create a wrapper div
        div.classList.add('image-wrapper', 'existing'); // Optional: Add a class for styling
        const img = document.createElement('img');
        img.src = `https://dxcportal.sharepoint.com${file.ServerRelativeUrl}`;
        const button = document.createElement('button'); // Create a button
        button.innerHTML = '<i class="fa-solid fa-xmark"></i>'; // Add the icon
        button.classList.add('remove-image-button');
        button.type = 'button';
        const restorebutton = document.createElement('button'); // Create a button
        restorebutton.innerHTML = '<i class="fa-solid fa-rotate-right"></i>'; // Add the icon
        restorebutton.classList.add('restore-image-button');
        restorebutton.type = 'button';
        // Create the overlay
        const overlay = document.createElement('div');
        overlay.classList.add('image-overlay');
        overlay.innerText = "Marked for deletion"; // Optional label
        // Add click event to remove the preview and update the file list
        button.addEventListener('click', function() {
            // temporarily store the file name in json
            let attachmentsToDelete = JSON.parse($('#attachmentsToDelete').val());
            if (!attachmentsToDelete.includes(file.FileName)) {
                attachmentsToDelete.push(file.FileName);
            }
            $('#attachmentsToDelete').val(JSON.stringify(attachmentsToDelete));
            // Show the overlay instead of removing the image
            overlay.style.display = "flex";
            restorebutton.style.display = "block"; // Show the restore button
            button.style.display = "none"; // Hide the remove button
            $(this).closest('.image-wrapper').addClass('marked-for-deletion');
        });
        restorebutton.addEventListener('click', function() {
            // temporarily store the file name in json
            let attachmentsToDelete = JSON.parse($('#attachmentsToDelete').val());
            //remove file name from list
            attachmentsToDelete = attachmentsToDelete.filter(item => item !== file.FileName);

            $('#attachmentsToDelete').val(JSON.stringify(attachmentsToDelete));
            overlay.style.display = "none"; //Hide the overlay
            restorebutton.style.display = "none"; // Hide the restore button
            button.style.display = "block"; // Show the remove button
            $(this).closest('.image-wrapper').removeClass('marked-for-deletion');
        });


        div.appendChild(img); // Append the image to the wrapper div
        div.appendChild(button); // Append the button to the wrapper div
        div.appendChild(restorebutton); // Append the button to the wrapper div
        div.appendChild(overlay); // Add overlay but keep it hidden initially
        previewContainer.appendChild(div); // Add the image to the preview container
    });

}
function filterEntryList(month,filter){
    switch(filter){
        case '1': //all entries
            getAllEntriesByMonth(month);
            break;
        case '2': //own entries
            getSelfEntriesByMonth(month);
            break;
        case '3': //own recognition
            getOwnRecognition(month);
            break;
    }
}
async function processListItems(items,month, sort = 0) {
    let target = $('#entriesContainer')//$('#'+month+'Entries');
    let html = ``;
    spItems = []; // clear spItems
    // Sort items based on sort
    switch(sort){
        case '1':
            items.sort((a, b) => new Date(a.Created) - new Date(b.Created)); // Sort by Created date ascending
            break;
        case '2':
            items.sort((a, b) => new Date(b.Created) - new Date(a.Created)); // Sort by Created date descending
            break;
        case '3':
            items.sort((a, b) => a.Submitted_x0020_By.localeCompare(b.Submitted_x0020_By)); // Sort by Submitted By ascending
            break;
        case '4':
            items.sort((a, b) => b.Submitted_x0020_By.localeCompare(a.Submitted_x0020_By)); // Sort by Submitted By descending
            break;
        case '5':
            items.sort((a, b) => monthNames.indexOf(a.Month) - monthNames.indexOf(b.Month)); // sort by month ascending
            break;
        case '6':
            items.sort((a, b) => monthNames.indexOf(b.Month) - monthNames.indexOf(a.Month)); // sort by month descending
            break;
        case '7':
            items.sort((a, b) => a.Account.localeCompare(b.Account));; // sort by account ascending
            break;
        case '8':
            items.sort((a, b) => b.Account.localeCompare(a.Account));; // sort by account ascending
            break;
        default:
            items.sort((a, b) => new Date(b.Created) - new Date(a.Created));
    }
    
    if(items.length > 0){
        for( const item of items){
            if(!item || !item.Id){ continue;}
            spItems[item.Id] = item;
            var uniqueness = item.Uniqueness ? item.Uniqueness : "N/A";
            var recognition = item.Individual ? "Individual" : "Team";
            var createdOn = new Date(`${item.Created}`);
            const options = { year: "numeric", month: "long", day: "numeric" };
            createdOn = createdOn.toLocaleDateString("en-US", options);
            var attachments = item.Attachments ? item.ID : null;
            let category = "Entry";
            if(access.type == 1){
                category = item.Amplified;
            }
            html += `<div class="card entry-card my-1 ${category} position-relative entry-preview" data-id="${item.Id}">`
            
            html+=`<div class="card-body">
                    <button type="button" title="Select" entry-id="${item.Id}" class="select-entry-button me-2" is-selected="0"><i class="fa-regular fa-square"></i></button>
                    <div style="max-width: 75%;">
                        <span class="lead" entry-id="${item.Id}" style="cursor: pointer">${item.Title}</span><br>
                        <span><b>${item.SUBSL}</b> | ${item.Account}`;
            if(month == 'All'){
                html += ` | ${item.Month}, ${item.Year}`;
            }
            html += ` | Entry by: ${item.Submitted_x0020_By}</span>`;
            html+=`</span>
                    </div>`
            html+= `<div class="entry-actions">`;
            
            html+=`<button type="button" title="View Info" data-id="${item.Id}" data-bs-toggle="offcanvas" data-bs-target="#entryInfoCanvas" aria-controls="entryInfoCanvas" class="entry-view"><i class="fa-solid fa-right-to-bracket"></i></button>`
            
            // if((authorId == item.AuthorId && access.type == 2) || access.type == 1){            
            //     html += `<button type="button" title="Edit" data-id="${item.Id}" data-bs-toggle="offcanvas" data-bs-target="#updateEntryCanvas" aria-controls="updateEntryCanvas" class="entry-update"><i class="fa-solid fa-pen"></i></button>
            //             <button class="entry-delete" title="Delete" data-id="${item.Id}"><i class="fa-solid fa-trash-can"></i></button>`;
            // }
            // html += `<button title="Generate slides" class="entry-generate" data-id="${item.Id}"><i class="fa-solid fa-file-powerpoint"></i></button>`;
            html +=       `</div>`; //close entry-actions
            html+=`</div>
            </div>`;
        };
    }else{
        html = noContentHTML();
    }
    
    target.html(html);
}

//DASHBOARD FUNCTIONS
function embedPowerBIReport() {
    const embedConfig = {
        type: "report",
        id: "<report-id>",
        embedUrl: "<embed-url>",
        accessToken: "<access-token>",
        settings: {
        panes: {
            filters: {
            visible: false, // Hides filters pane
            },
        },
        layoutType: models.LayoutType.Custom, // Required for custom layout
        navContentPaneEnabled: false, // Hides page navigation bar
        },
    };

    // Embed the Power BI report into the container
    const reportContainer = document.getElementById("reportContainer");
    const report = powerbi.embed(reportContainer, embedConfig);
}

// EVENTS
$(document).ready(function() {
    init();
    $(window).on('beforeinstallprompt', function (e) {
        e.preventDefault();         // Prevent automatic prompt
        deferredPrompt = e.originalEvent;         // Save the event
        
        $('#installBtn').show();    // Reveal install button
        console.log("This is installable");
    })
    .on('appinstalled', function () {
        console.log('App installed successfully');
        localStorage.setItem('isInstalled', 'true');
        $('#installBtn').hide();
    })
    // .on('contextmenu', function (e) {
    //     e.preventDefault();
    //     console.log("Right-click disabled");
    // })
    
    
    $(document)
    .on('click','#sendPrompt',function(){
        let text = $('#AIPrompt').val();
        callMyAI(text);
    })
    .on('click', '#signIn', function(){
        signIn();
    })
    .on('click', '#signOut', async function(){
        msalInstance.logout();
    })
    .on('change','#subslDropdown, #updateSubslDropdown',function() {
        if($(this).val() === ""){
            $(this)[0].setCustomValidity("Please select a valid option.");                
        } else {
            $(this)[0].setCustomValidity("");
        }
    })
    .on('change','#accountDropdown',function() {
        console.log($(this).val())
        if ($(this).val() === 'other') {
            $('#otherAccountContainer').show();
            $('#otherAccountInput').prop('required', true).attr('aria-required', 'true');
            $('#otherAccountInput').focus();
            $(this)[0].setCustomValidity("");
        }else if($(this).val() === ""){
            $(this)[0].setCustomValidity("Please select a valid option.");                
        } else {
            $(this)[0].setCustomValidity("");
            $('#otherAccountContainer').hide();
            $('#otherAccountInput').removeProp('required').removeAttr('aria-required').removeAttr('required');
        }
    })
    .on('change','#updateAccountDropdown',function() {
        if ($(this).val() === 'other') {
            $('#updateOtherAccountContainer').show();
            $('#updateOtherAccountInput').prop('required', true).attr('aria-required', 'true');
            $('#updateOtherAccountInput').focus();
            $(this)[0].setCustomValidity("");
        }else if($(this).val() === ""){
            $(this)[0].setCustomValidity("Please select a valid option.");                
        } else {
            $(this)[0].setCustomValidity("");
            $('#updateOtherAccountContainer').hide();
            $('#updateOtherAccountInput').removeProp('required').removeAttr('aria-required').removeAttr('required');
        }
    })
    .on('click mousedown','input[name="recognitionInput"], input[name="otherRecipientType"]',function(e) {
        e.preventDefault();
    })
    .on('click','.otherRecipientBtn',function(){
        $('.otherRecipientBtn').removeClass('clicked');
        $(this).addClass('clicked');
    })
    .on('click','.updateOtherRecipientBtn',function(){
        $('.updateOtherRecipientBtn').removeClass('clicked');
        $(this).addClass('clicked');
    })
    .on('click', '.removeRecipient', function() {
        $(this).parent().remove(); // Remove the recipient pill
        $('.recognitionPill').removeClass('bg-secondary-subtle');
        if($('.recipientPill').length === 0){
            $('#recipientInput').focus(); 
            $('#recipientImages').val("");
        }else if($('.recipientPill').length == 1 && !$('.recipientPill').hasClass('type-Team')){
            $('#recognitionInput1').prop('checked', true);
            $('#recognitionInput2').removeProp('checked').removeAttr('checked');
            $('.recognitionPill').removeClass('bg-secondary-subtle');
            $('#recognitionInput1Label').addClass('bg-secondary-subtle');
        }else if($('.recipientPill').length == 1 && $('.recipientPill').hasClass('type-Team')){
            $('#recognitionInput2').prop('checked', true);
            $('#recognitionInput1').removeProp('checked').removeAttr('checked');
            $('.recognitionPill').removeClass('bg-secondary-subtle');
            $('#recognitionInput2Label').addClass('bg-secondary-subtle');
        }else{
            $('#recognitionInput2').prop('checked', true);
            $('#recognitionInput1').removeProp('checked').removeAttr('checked');
            $('.recognitionPill').removeClass('bg-secondary-subtle');
            $('#recognitionInput2Label').addClass('bg-secondary-subtle');
        }
        document.getElementById('recipientImages').value = "";
        $('#recipientImages').replaceWith($('#recipientImages').clone());
    })
    .on('click', '.updateRemoveRecipient', function() {
        $(this).parent().remove(); // Remove the recipient pill
        $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
        if($('.updateRecipientPill').length === 0){
            $('#updateRecipientInput').focus(); 
            $('#updateRecipientImages').val("");
        }else if($('.updateRecipientPill').length == 1 && !$('.updateRecipientPill').hasClass('type-Team')){
            $('#updateRecognitionInput1').prop('checked', true);
            $('#updateRecognitionInput2').removeProp('checked').removeAttr('checked');
            $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
            $('#updateRecognitionInput1Label').addClass('bg-secondary-subtle');
        }else if($('.updateRecipientPill').length == 1 && $('.updateRecipientPill').hasClass('type-Team')){
            $('#updateRecognitionInput2').prop('checked', true);
            $('#updateRecognitionInput1').removeProp('checked').removeAttr('checked');
            $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
            $('#updateRecognitionInput2Label').addClass('bg-secondary-subtle');
        }else{
            $('#updateRecognitionInput2').prop('checked', true);
            $('#updateRecognitionInput1').removeProp('checked').removeAttr('checked');
            $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
            $('#updateRecognitionInput2Label').addClass('bg-secondary-subtle');
        }
        document.getElementById('updateRecipientImages').value = "";
        $('#updateRecipientImages').replaceWith($('#updateRecipientImages').clone());
    })
    .on('change','#recipientImages',function(event){
        const previewContainer = document.getElementById('previewContainer');
        let maxUploads = 0;
        let groupError = "";
        let recipients = $('.recipientPill').length;
        if(recipients <= 4 && recipients > 0){
            maxUploads = recipients; // Set maxUploads based on the number of recipients
        }else if( recipients > 4){
            maxUploads = 1; // Default to 1 if more than 4
            groupError = "For more than 4 recipients, it is recommended to upload <b>one</b> image in landscape orientation.";
        }
        console.log(`Max Uploads: ${maxUploads}`);
        // Clear previous previews
        previewContainer.innerHTML = '';
        const filesArray = Array.from(event.target.files); 

        // Check if the number of selected files exceeds the maximum allowed
        if(maxUploads === 0){
            $('#recipientInput').focus();
            document.getElementById('recipientImages').value = "";
            $('#recipientImages').replaceWith($('#recipientImages').clone());
            return;
        }
        if (filesArray.length > maxUploads) {
            document.getElementById('recipientImages').value = "";
            $('#recipientImages').replaceWith($('#recipientImages').clone());
            showModal("Notice", "You can only upload a maximum of "+maxUploads+" image(s). "+groupError);
            return;
        }
        for (let index = 0; index < filesArray.length; index++) {
            const file = filesArray[index];
            if (file.type.startsWith('image/')) {
                const reader = new FileReader();
                const image = new Image();

                reader.onload = function (e) {
                    image.onload = function () {
                        const width = image.width;
                        const height = image.height;
                        console.log(`File: ${file.name} → Width: ${width}, Height: ${height}`);

                        if (recipients > 4) {
                            if (width < height) {
                                groupError = "For more than 4 recipients, it is recommended to upload <b>one</b> image in landscape orientation.";
                                showModal("Notice", groupError);
                                return; // exits image.onload
                            } else if (width / height < 1.5 || width / height > 2) {
                                groupError = "For more than 4 recipients, it is recommended to upload <b>one</b> group image with an aspect ratio between 1.5:1 and 2:1. For example, <b>1920x1080</b> or <b>1600x900</b>.";
                                showModal("Notice", groupError);
                                return;
                            }
                            $('#previewContainer').append(`<div class="m-2" id="previewDimensions"><b>Width</b>: ${width}px<br><b>Height</b>: ${height}px</div>`);
                        }
                    };

                    const div = document.createElement('div');
                    div.classList.add('image-wrapper');
                    const img = document.createElement('img');
                    image.src = e.target.result;
                    img.src = e.target.result;
                    const button = document.createElement('button');
                    button.innerHTML = '<i class="fa-solid fa-xmark"></i>';
                    button.classList.add('remove-image-button');
                    button.type = 'button';

                    button.addEventListener('click', function () {
                        filesArray.splice(index, 1);
                        div.remove();
                        updateFileInput(filesArray, 'recipientImages');
                        $('#previewContainer').find('#previewDimensions').remove();
                    });

                    div.appendChild(img);
                    div.appendChild(button);
                    previewContainer.appendChild(div);
                };

                reader.readAsDataURL(file);
            }
        }

    })
    .on('change','#updateRecipientImages',function(event){
        const previewContainer = document.getElementById('updatePreviewContainer');
        let existingAttachments = $('.image-wrapper.existing:not(.marked-for-deletion)').length;
        const filesArray = Array.from(event.target.files); 
        let recipientImages = filesArray.length;
        let maxUploads = 0;
        let groupError = "";
        if($('.updateRecipientPill').length <= 4 && $('.updateRecipientPill').length > 0){
            maxUploads = $('.updateRecipientPill').length - existingAttachments; // Set maxUploads based on the number of recipients
        }else if($('.updateRecipientPill').length > 4){
            if(existingAttachments > 0)
                maxUploads = 0; // Default to 0 if more than 4 and existing attachments
            else{
                maxUploads = 1; // Default to 1 if more than 4
                groupError = "For more than 4 recipients, it is recommended to upload <b>one</b> image in landscape orientation.";
            }
                

        }

        // Clear previous previews
        previewContainer.innerHTML = '';

        // Check if the number of selected files exceeds the maximum allowed
        if(maxUploads === 0){
            document.getElementById('updateRecipientImages').value = "";
            $('#updateRecipientImages').replaceWith($('#updateRecipientImages').clone());
            if($('.updateRecipientPill').length > 0){
                $('#updateRecipientInput').focus();
                showModal("Error", `There are already ${existingAttachments} image(s) attached.`);
            }else{
                $('#updateRecipientInput').focus();
            }

            return;
        }
        if (filesArray.length > maxUploads) {
            document.getElementById('updateRecipientImages').value = "";
            $('#updateRecipientImages').replaceWith($('#updateRecipientImages').clone());
            showModal("Error", "You can only upload a maximum of "+maxUploads+" image(s). "+groupError);
            return;
        }
        
        // Loop through selected files
        filesArray.forEach((file, index) => {
            if (file.type.startsWith('image/')) { // Ensure the file is an image
                const reader = new FileReader();
                
                reader.onload = function(e) {
                    const div = document.createElement('div'); // Create a wrapper div
                    div.classList.add('image-wrapper'); // Optional: Add a class for styling
                    const img = document.createElement('img');
                    img.src = e.target.result; // Set the image source
                    const button = document.createElement('button'); // Create a button
                    button.innerHTML = '<i class="fa-solid fa-xmark"></i>'; // Add the icon
                    button.classList.add('remove-image-button');
                    button.type = 'button';
                    // Add click event to remove the preview and update the file list
                    button.addEventListener('click', function() {
                        filesArray.splice(index, 1); // Remove the file from the array
                        div.remove(); // Remove the div from the preview container
                        updateFileInput(filesArray,'updateRecipientImages'); // Update the fileInput element
                    });

                    div.appendChild(img); // Append the image to the wrapper div
                    div.appendChild(button); // Append the button to the wrapper div
                    previewContainer.appendChild(div); // Add the image to the preview container
                };
                
                reader.readAsDataURL(file); // Read the file as a data URL
            }
        });
    })
    .on('change','#typeOfWork',function() {        
        if ($(this).val() === 'others') {
            $('#otherTypeOfWorkContainer').show();
            $('#otherTypeOfWorkInput').prop('required', true).attr('aria-required', 'true');
            $('#otherTypeOfWorkInput').focus();
            $(this)[0].setCustomValidity("");
        }else if($(this).val() === ""){
            $(this)[0].setCustomValidity("Please select a valid option.");                
        } else {
            $(this)[0].setCustomValidity("");
            $('#otherTypeOfWorkContainer').hide();
            $('#otherTypeOfWorkInput').removeProp('required').removeAttr('aria-required').removeAttr('required');
        }
    })
    .on('change','#updateTypeOfWork',function() {
        if ($(this).val() === 'others') {
            $('#updateOtherTypeOfWorkContainer').show();
            $('#updateOtherTypeOfWorkInput').prop('required', true).attr('aria-required', 'true');
            $('#updateOtherTypeOfWorkInput').focus();
            $(this)[0].setCustomValidity("");
        }else if($(this).val() === ""){
            $(this)[0].setCustomValidity("Please select a valid option.");                
        } else {
            $(this)[0].setCustomValidity("");
            $('#updateOtherTypeOfWorkContainer').hide();
            $('#updateOtherTypeOfWorkInput').removeProp('required').removeAttr('aria-required').removeAttr('required');
        }
    })
    .on('change','#uniquenessTickbox',function() {
        if ($(this).is(':checked')) {
            $('#uniquenessInput').show().prop('required', true).attr('aria-required', 'true');
            $('#uniquenessInput').next('.text-count').show();
            $('#uniquenessInput').focus();
        } else {
            $('#uniquenessInput').hide().removeProp('required').removeAttr('aria-required').removeAttr('required');
            $('#uniquenessInput').next('.text-count').hide();
        }
    })
    .on('change','#updateUniquenessTickbox',function() {
        if ($(this).is(':checked')) {
            $('#updateUniquenessInput').show().prop('required', true).attr('aria-required', 'true');
            $('#updateUniquenessInput').focus();
        } else {
            $('#updateUniquenessInput').hide().removeProp('required').removeAttr('aria-required').removeAttr('required');
        }
    })
    .on('input','#recipientInput', async function(event){
        const searchTerm = event.target.value.toLowerCase();

        if (searchTerm.length >= 2) { // Fetch users only if input is 2+ characters
            const users = await fetchTenantUsers(searchTerm);

            // Display suggestions in the dropdown
            const dropdown = document.getElementById("recipientDropdown");
            dropdown.innerHTML = ""; // Clear previous suggestions
            users.forEach(user => {
                const option = document.createElement("div");
                option.textContent = `${user.name} (${user.email})`;
                option.classList.add("recipient-dropdown-item");
                option.setAttribute("data-user-id",user.id);
                option.setAttribute("data-name",user.name);
                option.setAttribute("data-email",user.email);
                dropdown.appendChild(option);
            });
            $('#recipientDropdown').show();
        }
    })
    .on('input','#updateRecipientInput', async function(event){
        const searchTerm = event.target.value.toLowerCase();

        if (searchTerm.length >= 2) { // Fetch users only if input is 2+ characters
            const users = await fetchTenantUsers(searchTerm);

            // Display suggestions in the dropdown
            const dropdown = document.getElementById("updateRecipientDropdown");
            dropdown.innerHTML = ""; // Clear previous suggestions
            users.forEach(user => {
                const option = document.createElement("div");
                option.textContent = `${user.name} (${user.email})`;
                option.classList.add("update-recipient-dropdown-item");
                option.setAttribute("data-user-id",user.id);
                option.setAttribute("data-name",user.name);
                option.setAttribute("data-email",user.email);
                dropdown.appendChild(option);
            });
            $('#updateRecipientDropdown').show();
        }
    })
    .on('click','#otherRecipientOption',function(){
        $('#otherRecipientContainer').toggleClass('d-none d-flex');
        $('#recipientDropdown').empty().hide();
        $('#otherRecipientInput').val('');
        $('#otherRecipientEmailInput').attr('value', 'no-email@dxc.com');
    })
    .on('click','#updateOtherRecipientOption',function(){
        $('#updateOtherRecipientContainer').toggleClass('d-none d-flex');
        $('#updateRecipientDropdown').empty().hide();
        $('#updateOtherRecipientInput').val('');
        $('#updateOtherRecipientEmailInput').attr('value', 'no-email@dxc.com');
    })
    .on('click', '#addOtherRecipient',function(){
        const otherRecipientInput = $('#otherRecipientInput');
        const otherRecipientEmailInput = $('#otherRecipientEmailInput');
        const otherRecipientName = otherRecipientInput.val().trim();
        const otherRecipientEmail = otherRecipientEmailInput.val().trim();
        if(otherRecipientName === '' || otherRecipientEmail === ''){
            showModal("Error", "Please enter a name and email for the other recipient.");
            return;
        }
        const recipientType = "type-"+$('.otherRecipientBtn.clicked').attr('data-type');
        const recipientContainer = $('#recipientContainer');
        const newRecipient = $(`<div class="recipientPill bg-secondary-subtle px-2 ${recipientType}" value="${otherRecipientEmail}" name="${otherRecipientName}">${otherRecipientName}<button class="removeRecipient" type="button"><i class="fa-solid fa-xmark"></i></button></div>`);
        recipientContainer.append(newRecipient);
        otherRecipientInput.val(''); // Clear the input field
        otherRecipientEmailInput.val(''); // Clear the email field
        $('#otherRecipientContainer').removeClass('d-flex').addClass('d-none'); // Hide the other recipient container
        $('#recipientDropdown').empty().hide(); // Clear the dropdown suggestions
        if($('.recipientPill').length > 0){
            if($('.recipientPill').length == 1 && !$('.recipientPill').hasClass('type-Team')){
                $('#recognitionInput1').prop('checked', true);
                $('#recognitionInput2').removeProp('checked').removeAttr('checked');
                $('.recognitionPill').removeClass('bg-secondary-subtle');
                $('#recognitionInput1Label').addClass('bg-secondary-subtle');
            }else if($('.recipientPill').length == 1 && $('.recipientPill').hasClass('type-Team')){
                $('#recognitionInput2').prop('checked', true);
                $('#recognitionInput1').removeProp('checked').removeAttr('checked');
                $('.recognitionPill').removeClass('bg-secondary-subtle');
                $('#recognitionInput2Label').addClass('bg-secondary-subtle');
            }else{
                $('#recognitionInput2').prop('checked', true);
                $('#recognitionInput1').removeProp('checked').removeAttr('checked');
                $('.recognitionPill').removeClass('bg-secondary-subtle');
                $('#recognitionInput2Label').addClass('bg-secondary-subtle');
            }
        }
    })
    .on('click', '#updateAddOtherRecipient',function(){
        const otherRecipientInput = $('#updateOtherRecipientInput');
        const otherRecipientEmailInput = $('#updateOtherRecipientEmailInput');
        const otherRecipientName = otherRecipientInput.val().trim();
        const otherRecipientEmail = otherRecipientEmailInput.val().trim();
        if(otherRecipientName === '' || otherRecipientEmail === ''){
            showModal("Error", "Please enter a name and email for the other recipient.");
            return;
        }
        const recipientType = "type-"+$('.updateOtherRecipientBtn.clicked').attr('data-type');
        const recipientContainer = $('#updateRecipientContainer');
        const newRecipient = $(`<div class="updateRecipientPill bg-secondary-subtle px-2 ${recipientType}" value="${otherRecipientEmail}" name="${otherRecipientName}">${otherRecipientName}<button class="updateRemoveRecipient" type="button"><i class="fa-solid fa-xmark"></i></button></div>`);
        recipientContainer.append(newRecipient);
        otherRecipientInput.val(''); // Clear the input field
        otherRecipientEmailInput.val(''); // Clear the email field
        $('#updateOtherRecipientContainer').removeClass('d-flex').addClass('d-none'); // Hide the other recipient container
        $('#updateRecipientDropdown').empty().hide(); // Clear the dropdown suggestions
        if($('.updateRecipientPill').length > 0){
            if($('.updateRecipientPill').length == 1 && !$('.updateRecipientPill').hasClass('type-Team')){
                $('#updateRecognitionInput1').prop('checked', true);
                $('#updateRecognitionInput2').removeProp('checked').removeAttr('checked');
                $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
                $('#updateRecognitionInput1Label').addClass('bg-secondary-subtle');
            }else if($('.updateRecipientPill').length == 1 && $('.updateRecipientPill').hasClass('type-Team')){
                $('#updateRecognitionInput2').prop('checked', true);
                $('#updateRecognitionInput1').removeProp('checked').removeAttr('checked');
                $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
                $('#updateRecognitionInput2Label').addClass('bg-secondary-subtle');
            }else{
                $('#updateRecognitionInput2').prop('checked', true);
                $('#updateRecognitionInput1').removeProp('checked').removeAttr('checked');
                $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
                $('#updateRecognitionInput2Label').addClass('bg-secondary-subtle');
            }
        }
    })
    .on('click','.recipient-dropdown-item',function(){
        const selectedUser = $(this).data('name');
        const userId = $(this).data('user-id');
        const email = $(this).data('email');
        const recipientContainer = $('#recipientContainer');
        const newRecipient = $(`<div class="recipientPill bg-secondary-subtle px-2" value="${email}" name="${selectedUser}">${selectedUser}<button class="removeRecipient" type="button"><i class="fa-solid fa-xmark"></i></button></div>`);
        recipientContainer.append(newRecipient);
        $('#recipientInput').val(''); // Clear the input field
        $('#recipientDropdown').empty().hide(); // Clear the dropdown suggestions
        if($('.recipientPill').length > 0){
            if($('.recipientPill').length == 1){
                $('#recognitionInput1').prop('checked', true);
                $('#recognitionInput2').removeProp('checked').removeAttr('checked');
                $('.recognitionPill').removeClass('bg-secondary-subtle');
                $('#recognitionInput1Label').addClass('bg-secondary-subtle');
            }else{
                $('#recognitionInput2').prop('checked', true);
                $('#recognitionInput1').removeProp('checked').removeAttr('checked');
                $('.recognitionPill').removeClass('bg-secondary-subtle');
                $('#recognitionInput2Label').addClass('bg-secondary-subtle');
            }
            document.getElementById('recipientImages').value = "";
            $('#recipientImages').replaceWith($('#recipientImages').clone());// Clear the file input
            $('#previewContainer').empty();
        }
    })
    .on('click','.update-recipient-dropdown-item',function(){
        const selectedUser = $(this).data('name');
        const userId = $(this).data('user-id');
        const email = $(this).data('email');
        const recipientContainer = $('#updateRecipientContainer');
        const newRecipient = $(`<div class="updateRecipientPill bg-secondary-subtle px-2" value="${email}" name="${selectedUser}">${selectedUser}<button class="updateRemoveRecipient" type="button"><i class="fa-solid fa-xmark"></i></button></div>`);
        recipientContainer.append(newRecipient);
        $('#updateRecipientInput').val(''); // Clear the input field
        $('#updateRecipientDropdown').empty().hide(); // Clear the dropdown suggestions
        if($('.updateRecipientPill').length > 0){
            if($('.updateRecipientPill').length == 1){
                $('#updateRecognitionInput1').prop('checked', true);
                $('#updateRecognitionInput2').removeProp('checked').removeAttr('checked');
                $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
                $('#updateRecognitionInput1Label').addClass('bg-secondary-subtle');
            }else{
                $('#updateRecognitionInput2').prop('checked', true);
                $('#updateRecognitionInput1').removeProp('checked').removeAttr('checked');
                $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
                $('#updateRecognitionInput2Label').addClass('bg-secondary-subtle');
            }
            $('#recipientImages').val(""); // Clear the file input
            $('#updatePreviewContainer').empty();
        }
    })
    .on('input', '#businessChallengeInput, #businessImpactInput, #uniquenessInput ,#howDXCHelpedInput',function(){
        let length = $(this).val().trim().length;
        $(this).next('.text-count').text(length)
    })
    .on('focus', '#businessChallengeInput, #businessImpactInput, #uniquenessInput ,#howDXCHelpedInput',function(){
        let id = $(this).attr('id');
        $('.mistral-button').removeClass('active')
        $('.mistral-button[textarea='+id+']').addClass('active');
    })
    .on('submit','#newEntryForm', async function(e){
        e.preventDefault();
        if($('#accountDropdown').val() == ""){
            $('#accountDropdown').setCustomValidity("Please select a valid option.");
            $('#accountDropdown').focus();
            console.log("account dropdown value == empty");
            return false;
        }else if ($('#subslDropdown').val() == ""){
            $('#subslDropdown').setCustomValidity("Please select a valid option.");
            $('#subslDropdown').focus();
            console.log("subslDropdown value == empty");
            return false;
        }else if ($('#typeOfWork').val() == ""){
            $('#typeOfWork').setCustomValidity("Please select a valid option.");
            $('#typeOfWork').focus();
            console.log("typeOfWork value == empty");
            return false;
        }
        let submit = $(this).find('button[type="submit"]');
        spinButton(submit);
        let list = "Amazing Stories entries dev";
        let data = {};
        let attachments = [];
        let form = $(this);
        let year = new Date().getFullYear().toString();
        let month = $('#monthDropdown').val();
        let subsl = $('#subslDropdown').val();
        let account = $('#accountDropdown').val() === 'other' ? $('#otherAccountInput').val().trim() : $('#accountDropdown').val();
        let team = $('#teamInput').val().trim();
        let recognition = $('input[name="recognitionInput"]:checked').attr('id') === 'recognitionInput1' ? true : false; //true for individual, false for team
        let entry = $('#entryInput').val().trim();//title field
        let recipients = [];
        let recipientEmails = [];

        if($('.recipientPill').length > 0){
            $('.recipientPill').each(function(){
                recipientEmails.push($(this).attr('value'));
                recipients.push($(this).attr('name').trim());
            });
        }else{
            showModal("Error", "Please add at least one recipient.");
            resetFormValidation(form);
            stopSpinButton(submit, "Submit");
            return;
        }
        let typeOfWork = $('#typeOfWork').val() === 'others' ? $('#otherTypeOfWorkInput').val().trim() : $('#typeOfWork').val();
        let businessChallenge = $('#businessChallengeInput').val().trim();
        let howDXCHelped = $('#howDXCHelpedInput').val().trim();
        let businessImpact = $('#businessImpactInput').val().trim();
        let uniqueness = $('#uniquenessTickbox').is(':checked') ? $('#uniquenessInput').val().trim() : null;
        let outcome = $('#outcomeTickbox').is(':checked') ? $('#outcomeInput').val().trim() : null;
        let recipientImages = $('#recipientImages')[0].files;
        
        if (recipientImages.length > 0) {
            // Iterate through each uploaded file
            let imageData = new FormData();
            Array.from(recipientImages).forEach(file => {
                let fileName = file.name.split(".")[0];  // "cat2"
                let fileExtension = file.name.split(".").pop();  // "png"
                let newFileName = fileName + getRandomString(5) +"." + fileExtension;  // "cat2abcde.png"
                
                // Create a new File object with the renamed filename
                let renamedFile = new File([file], newFileName, { type: file.type });
                
                attachments.push(renamedFile); // Add to attachments array
                imageData.append("images", renamedFile); // Append renamed file to FormData
                
                console.log(`File added: ${renamedFile.name} (${renamedFile.type}, ${renamedFile.size} bytes)`);
            });
            // const response = await fetch("/api/upload", {
            //     method: "POST",
            //     body: imageData
            // });

            // const result = await response.json();
            // console.log("Uploaded:", result);
        } else {
            console.log("No files uploaded.");
        }
        data = {
            "Title": entry,
            "Year": year,
            "Month": month,
            "SUBSL": subsl,
            "Account": account,
            "Team": team,
            "Individual": recognition,
            "Recipients": recipients.join('/ '),
            "Recipient_x0020_Emails": recipientEmails.join('/ '),
            "Worktype": typeOfWork,
            "Challenge": businessChallenge,
            "Help": howDXCHelped,
            "Impact": businessImpact,
            "Uniqueness": uniqueness,
            "Outcome": outcome,
            "Amplified": "Entry",
            "Submitted_x0020_By": email,
        };
        if(access.type != 3){
            data["Classification"] = $('#classificationSelect').val();
        }
        saveToSharePoint(splist, data, attachments).then(()=>{
            showModal("Success", "Entry submitted successfully.");
            $('#newEntryForm')[0].reset(); // Reset the form fields
            $('#recipientContainer').empty(); // Clear the recipient container
            $('#previewContainer').empty(); // Clear the preview container
            $('#recipientImages').val(""); // Clear the file input
            setCurrentMonth(); // Reset the month dropdown to the current month
            $('#uniquenessInput').hide().removeProp('required').removeAttr('aria-required').removeAttr('required');
            $('#uniquenessInput').next('.text-count').hide();
            updateFileInput([],'recipientImages'); // Clear the file input's FileList
            $('#otherTypeOfWorkContainer').hide()
            $('#otherTypeOfWorkInput').removeProp('required').removeAttr('aria-required').removeAttr('required');
        }).catch((error) => {
            console.error("Error saving to SharePoint:", error);
            showModal("Error", `An error occurred while submitting the entry. Please try again. <hr><br>${error.message}`);
        })

        .finally(() => {
            resetFormValidation(form);
            stopSpinButton(submit, "Submit");
        });

    })
    .on('submit','#updateEntryForm',function(e){
        e.preventDefault();
        if($('#updateAccountDropdown').val() === ""){
            $('#updateAccountDropdown').setCustomValidity("Please select a valid option.");
            $('#updateAccountDropdown').focus();
            return false;
        }else if ($('#updateSubslDropdown').val() === ""){
            $('#updateSubslDropdown').setCustomValidity("Please select a valid option.");
            $('#updateSubslDropdown').focus();
            return false;
        }else if ($('#updateTypeOfWork').val() === ""){
            $('#updateTypeOfWork').setCustomValidity("Please select a valid option.");
            $('#updateTypeOfWork').focus();
            return false;
        }
        if ($("#updateOtherAccountInput").is(":visible")) {

        }

        let submit = $(this).find('button[type="submit"]');
        spinButton(submit);
        let id = $('#updateId').val();
        let list = "Amazing Stories entries";
        let data = {};
        let attachments = [];
        let form = $(this);
        let year = $('#updateYearInput').val();
        let month = $('#updateMonthDropdown').val();
        let subsl = $('#updateSubslDropdown').val();
        let account = $('#updateAccountDropdown').val() === 'other' ? $('#updateOtherAccountInput').val().trim() : $('#updateAccountDropdown').val();
        let team = $('#updateTeamInput').val().trim();
        let recognition = $('input[name="updateRecognitionInput"]:checked').attr('id') === 'updateRecognitionInput1' ? true : false; //true for individual, false for team
        let entry = $('#updateEntryInput').val().trim(); // title field
        let recipients = [];
        let recipientEmails = [];
        let attachmentsToDelete = JSON.parse($('#attachmentsToDelete').val());
        if ($('.updateRecipientPill').length > 0) {
            $('.updateRecipientPill').each(function() {
                recipientEmails.push($(this).attr('value'));
                recipients.push($(this).attr('name').trim());
            });
        } else {
            showModal("Error", "Please add at least one recipient.");
            resetFormValidation(form);
            stopSpinButton(submit, "Submit");
            return;
        }
        let existingAttachments = $('.image-wrapper.existing').not('.marked-for-deletion').length;
        let recipientImages = $('#updateRecipientImages')[0].files;
        let typeOfWork = $('#updateTypeOfWork').val() === 'others' ? $('#updateOtherTypeOfWorkInput').val().trim() : $('#updateTypeOfWork').val();
        let businessChallenge = $('#updateBusinessChallengeInput').val().trim();
        let howDXCHelped = $('#updateHowDXCHelpedInput').val().trim();
        let businessImpact = $('#updateBusinessImpactInput').val().trim();
        let uniqueness = $('#updateUniquenessTickbox').is(':checked') ? $('#updateUniquenessInput').val().trim() : null;

        //check against number of recipients
        let totalAttachments = existingAttachments + recipientImages.length;
        let maxAttachments = $('.updateRecipientPill').length <= 4 ? $('.updateRecipientPill').length : 4;
        console.log("Total attachments: " + totalAttachments);
        if (totalAttachments > maxAttachments) {
            showModal("Error", "You can only upload a maximum of <b>"+maxAttachments+"</b> image/s (one per recipient but maximum of 4 in total).");
            resetFormValidation(form);
            stopSpinButton(submit, "Submit");
            $('#updateRecipientImages').val("");
            return false;
        }

        // Iterate through each uploaded file
        Array.from(recipientImages).forEach(file => {
            attachments.push(file); // Add each file to the attachments array
            console.log(`File added: ${file.name} (${file.type}, ${file.size} bytes)`);
        });
        //compare fields
        let entryItem = spItems[id];
        
        let fields = [
            {name:"Title", value: entry},
            {name:"Year", value: year},
            {name:"Month", value: month},
            {name:"SUBSL", value: subsl},
            {name:"Account", value: account},
            {name:"Team", value: team},
            {name:"Individual", value: recognition},
            {name:"Recipients", value: recipients.join('/ ')},
            {name:"Recipient_x0020_Emails", value: recipientEmails.join('/ ')},
            {name:"Worktype", value: typeOfWork},
            {name:"Challenge", value: businessChallenge},
            {name:"Help", value: howDXCHelped},
            {name:"Impact", value: businessImpact},
            {name:"Uniqueness", value: uniqueness},
        ];
        if(access.type != 3){
            let classification = $('#updateClassificationSelect').val();
            fields.push({name:"Classification", value: classification});
        }
        let filteredFields = fields.filter(field => entryItem[field.name] !== field.value);

        // return false;
        updateSPItem(splist, id, filteredFields).then(async ()=>{
            if(attachments.length > 0){
                await addAttachments(id, attachments).then(()=>{
                    console.log("Attachments saved successfully.");
                }).catch((error) => {
                    console.error("Error saving attachments:", error);
                });
            }
            if(attachmentsToDelete.length > 0){
                await deleteAttachments(splist, id, attachmentsToDelete).then(()=>{
                    console.log("Attachments deleted successfully.");
                    attachmentsToDelete.forEach(attachment => {
                        // Remove the corresponding image from the preview container
                        $('#entryInfoImages img[data-name="'+attachment+'"]').remove();
                    });
                    if($('#entryInfoImages img').length == 0){
                        $('#entryInfoImages').html(`<p>No attachments.</p>`)
                    }
                }).catch((error) => {
                    console.error("Error deleting attachments:", error);
                });
            }
            if(attachments.length > 0 || attachmentsToDelete.length > 0){
                filteredFields.push({name: "Attachments", value: true});
            }
            createModifyEmail(spItems[id],filteredFields);
            let updated = await getEntryById(id);
            spItems[id] = updated;
            showModal("Success", "Entry updated successfully.");
            $('#updateEntryForm')[0].reset(); // Reset the form fields
            $('#updateRecipientContainer').empty(); // Clear the recipient container
            $('#updatePreviewContainer').empty(); // Clear the preview container
            $('#updateRecipientImages').val(""); // Clear the file input
            $('#updateOtherAccountInput').hide().removeProp('required').removeAttr('aria-required').removeAttr('required');
            $('#updateOtherTypeOfWorkInput').removeProp('required').removeAttr('aria-required').removeAttr('required');
            $('#updateOtherTypeOfWorkContainer').hide()
            $('#updateUniquenessInput').hide().removeProp('required').removeAttr('aria-required').removeAttr('required');
            updateFileInput([],'updateRecipientImages'); // Clear the file input's FileList
            $('#attachmentsToDelete').val("[]");
            viewItem(id);
            $('#updateEntryCanvasClose').click();
            // let labelId = $('input[name="entryFilter"]:checked').val();
            // $('.entryFilterLabel[value="'+labelId+'"]').click();
        }).catch((error) => {
            console.error("Error saving to SharePoint:", error);
            showModal("Error", "An error occurred while submitting the entry. Please try again.");
        })

        .finally(() => {
            resetFormValidation(form);
            stopSpinButton(submit, "Submit");
        });

    })
    .on('click','#openGenerateSlide',function(){
        let month = $('#monthFilter').val();
        var filter = $('input[name="entryFilter"]:checked').val();
        if(access.type == 3){filter = '3';}
        filterEntryList(month,filter)
    })
    .on('change','#subslFilter',function(){
        let sort = $('#sorter').val();
        let month = $('#monthFilter').val();
        var filter = $('input[name="entryFilter"]:checked').val();
        if(access.type == 3){return false;}
        filterEntryList(month,filter)
    })
    .on('change','#monthFilter',function(){
        let sort = $('#sorter').val();
        let month = $(this).val();
        var filter = $('input[name="entryFilter"]:checked').val();
        if(access.type == 3){filter = '3';}
        if(filter != 'All'){$('.month-sort').hide();}
        else{$('.month-sort').show();}
        filterEntryList(month,filter,sort);
    })
    .on('click','.entryFilterLabel',function(){
        let sort = $('#sorter').val();
        //get active month tab
        let month = $('#monthFilter').val();
        //get filter
        var filter = $(this).attr('value');
        if(access.type == 3){filter = '3';}
        if (filter == '3' && access.type != 3){
            $('#subslFilter').addClass('d-none');
        }else{
            $('#subslFilter').removeClass('d-none');
        }
        filterEntryList(month,filter);
        $('#entryPreviewPlaceholder').removeClass('d-none');
        $('#entryPreview').addClass('d-none');
    })
    .on('click','.select-entry-button', function(){
        let entryId = $(this).attr('entry-id');
        let isSelected = $(this).attr('is-selected');
        selectedEntries = [];
        if(isSelected == "0"){
            $(this).attr('is-selected', '1');
            $(this).find('i').removeClass('fa-regular fa-square').addClass('fa-solid fa-square-check');
        }else{
            $(this).attr('is-selected', '0');
            $(this).find('i').removeClass('fa-solid fa-square-check').addClass('fa-regular fa-square');
        }
        $('.select-entry-button[is-selected="1"]').each(function(){
            let id=$(this).attr('entry-id');
            selectedEntries.push(spItems[id]);
            
        });

        $('#selectedEntriesCount').text(selectedEntries.length);
    })
    .on('click','.entry-view', async function(){
        $('#entryInfoCard').hide();
        let entry = $(this).closest('.entry-card');
        let id = entry.data('id');
        viewItem(id);
    })
    .on('click','.entry-preview', function(){
        let id = $(this).data('id');
        let entryItem = spItems[id];
        let title = entryItem.Title;
        let year = entryItem.Year;
        let month = entryItem.Month;
        let subsl = entryItem.SUBSL;
        let account = entryItem.Account;
        let team = entryItem.Team;
        let recognition = entryItem.Individual ? "Individual" : "Team";
        let recipients = entryItem.Recipients.split('/ ');
        recipients = recipients.join(', ');
        let submittedby = entryItem.Submitted_x0020_By;
        let created = new Date(entryItem.Created);
        let typeOfWork = entryItem.Worktype;
        let challenge = entryItem.Challenge;
        let help = entryItem.Help;
        let impact = entryItem.Impact;
        let uniqueness = entryItem.Uniqueness && entryItem.Uniqueness.trim().length > 0 ? entryItem.Uniqueness : "None";
        $('#entryPreviewTitle').text(title);
        $('#entryPreviewYear').text(year);
        $('#entryPreviewMonth').text(month);
        $('#entryPreviewSubsl').text(subsl);
        $('#entryPreviewAccount').text(account);
        $('#entryPreviewTeam').text(team);
        $('#entryPreviewRecognition').text(recognition);
        $('#entryPreviewRecipients').text(recipients);
        $('#entryPreviewSubmittedBy').text(submittedby);
        $('#entryPreviewSubmittedOn').text(created.toLocaleDateString());
        $('#entryPreviewWorkType').text(typeOfWork);
        $('#entryPreviewChallenge').text(challenge);
        $('#entryPreviewHelp').text(help);
        $('#entryPreviewImpact').text(impact);
        $('#entryPreviewUniqueness').text(uniqueness);
        $('#entryPreview').removeClass('d-none');
        $('#entryPreviewPlaceholder').addClass('d-none')
    })
    .on('click','#EntryInfoClose',function(){
        $('#entryLoading').removeClass('d-none').addClass('d-flex');
        $('#entryInfoCard').hide();
        $('#entryInfoImages').html("");
        $('#entryInfoAmplifyBtn').removeClass('text-theme');
        let closebtn = $('#updateEntryCanvasClose');
        closebtn.attr('data-bs-target','#generateSlidesCanvas').attr('aria-controls','generateSlidesCanvas');
    })
    .on('click','#entryInfoEditBtn', function(){
        let closebtn = $('#updateEntryCanvasClose');
        closebtn.attr('data-bs-target','#entryInfoCanvas').attr('aria-controls','entryInfoCanvas');
    })
    .on('click', '.entry-update', function() {
        $('#attachmentsToDelete').val("[]")
        let id = $(this).attr('data-id');
        let entryItem = spItems[id];
        let title = entryItem.Title;
        let year = entryItem.Year;
        let month = entryItem.Month;
        let subsl = entryItem.SUBSL;
        let account = entryItem.Account;
        let team = entryItem.Team;
        let recognition = entryItem.Individual ? "Individual" : "Team";
        let recipients = entryItem.Recipients.split('/ ');
        let recipientEmails = entryItem.Recipient_x0020_Emails.split('/ ');
        let recipientHTML = "";
        let worktype = entryItem.Worktype;
        let challenge = entryItem.Challenge;
        let help = entryItem.Help;
        let impact = entryItem.Impact;
        let uniqueness = entryItem.Uniqueness && entryItem.Uniqueness.trim().length > 0 ? entryItem.Uniqueness : "";
        let classification = entryItem.Classification ? entryItem.Classification : "C&I PH Commendation";
        // let author = await getUserDetailsById(entry.data('author'));
        // let createdBy = author ? author.Title : "Unknown";
        $('#updateId').val(id);
        $('#updateEntryInput').val(title);
        $('#updateYearInput').val(year);
        $('#updateMonthDropdown').val(month);
        $('#updateSubslDropdown').val(subsl);
        $('#updateClassificationSelect').val(classification);
        //check if account exists in dropdown
        if($('#updateAccountDropdown option[value="'+account+'"]').length > 0){
            $('#updateAccountDropdown').val(account);
        }else{
            $('#updateAccountDropdown').val('other');
            $('#updateOtherAccountInput').val(account).show().prop('required', true).attr('aria-required', 'true').focus();
            $('#updateOtherAccountContainer').show();
            $('#updateOtherAccountInput').focus();
        }
        $('#updateTeamInput').val(team);
        //make recipient pills per recipient
        $('#updateRecipientContainer').empty(); // Clear the recipient container
        recipients.forEach((recipient,index) => {
            const newRecipient = $(`<div class="updateRecipientPill bg-secondary-subtle px-2" value="${recipientEmails[index]}" name="${recipient}">${recipient}<button class="updateRemoveRecipient" type="button"><i class="fa-solid fa-xmark"></i></button></div>`);
            $('#updateRecipientContainer').append(newRecipient);
        });
        $('#updateRecipientDropdown').empty().hide(); // Clear the dropdown suggestions
        if(recipients.length > 0){
            if(recipients.length == 1 && entryItem.Individual){
                $('#updateRecipientContainer .updateRecipientPill').addClass('type-Team')
            }
            if(entryItem.Individual){
                $('#updateRecognitionInput1').prop('checked', true);
                $('#updateRecognitionInput2').removeProp('checked').removeAttr('checked');
                $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
                $('#updateRecognitionInput1Label').addClass('bg-secondary-subtle');
            }else{
                $('#updateRecognitionInput2').prop('checked', true);
                $('#updateRecognitionInput1').removeProp('checked').removeAttr('checked');
                $('.updateRecognitionPill').removeClass('bg-secondary-subtle');
                $('#updateRecognitionInput2Label').addClass('bg-secondary-subtle');
            }
            $('#recipientImages').val(""); // Clear the file input
        }
        //display preview of attachments
        $('#updatePreviewContainer').empty();
        $('#existingPreviewContainer').empty();
        displayExistingAttachmentsInUpdate(id)
        //check if worktype is others
        if($('#updateTypeOfWork option[value="'+worktype+'"]').length > 0){
            $('#updateTypeOfWork').val(worktype);
            $('#updateOtherTypeOfWorkContainer').hide();
            $('#updateOtherTypeOfWorkInput').removeProp('required').removeAttr('aria-required').removeAttr('required').val('');
        }else {
            $('#updateTypeOfWork').val('others');
            $('#updateOtherTypeOfWorkContainer').show();
            $('#updateOtherTypeOfWorkInput').prop('required', true).attr('aria-required', 'true').val(worktype);
        }
            
        $('#updateBusinessChallengeInput').val(challenge);
        $('#updateHowDXCHelpedInput').val(help);
        $('#updateBusinessImpactInput').val(impact);
        //check if uniqueness has value
        if(uniqueness && uniqueness.trim().length > 0 && uniqueness != "N/A"){
            $('#updateUniquenessTickbox').prop('checked', true);
            $('#updateUniquenessInput').show().prop('required', true).attr('aria-required', 'true').val(uniqueness);
        }

        
    })
    .on('click','.entry-delete', function(){
        let id = $(this).attr('data-id');
        let entry = $('.entry-card[data-id="'+id+'"]');
        let entryItem = spItems[id];
        let title = entryItem.Title;
        let localTime = new Date();
        let utcOffset = 8 * 60 * 60 * 1000; // Convert GMT+8 offset to milliseconds
        let adjustedTime = new Date(localTime.getTime() + utcOffset).toISOString();
        let fields = [
            {name: "Is_x0020_Deleted", value: true},
            {name: "Deleted_x0020_By", value: email},
            {name: "Deleted", value: adjustedTime}
        ];
        showConfirmModal(`Are you sure you want to delete the entry <b>${title}</b>?`,()=>{
            updateSPItem(splist, id, fields).then(()=>{
                showModal("Success", "Entry deleted successfully.");
                entry.remove();
                createDeleteEmail(title,"deleted");
            }).catch((error) => {
                console.error("Error deleting entry:", error);
                showModal("Error", "An error occurred while deleting the entry. Please try again.");
            })
        })
        
    })
    .on('click','.entry-amplify',function(){
        let id = $(this).attr('data-id');
        let entry = $('.entry-card[data-id="'+id+'"]');
        let entryItem = spItems[id];
        let title = entryItem.Title;
        let amplified = $(this).attr('data-amplified');
        let color = "";
        let updateValue = "Entry";
        let statusMessage = ["nominate", "amplify", "unamplify"];
        let statusMessageIndex = 2;
        switch(amplified){
            case "Entry":
                updateValue = "Candidate";
                statusMessageIndex = 0;
                color = "text-theme2";
                break;
            case "Candidate":
                updateValue = "Amplified";
                statusMessageIndex = 1;
                color = "text-theme";
                break;
            case "Amplified":
                updateValue = "Entry";
                statusMessageIndex = 2;
                break;
        }
        
        let fields = [{name: "Amplified", value: updateValue}];
        $(this).attr('title', `Click to ${statusMessage[statusMessageIndex]} this entry.`)
        showConfirmModal(`Are you sure you want to ${statusMessage[statusMessageIndex]} the entry <b>${title}</b>?`,()=>{
            updateSPItem(splist, id, fields).then(()=>{
                showModal("Success", `Entry updated successfully.`);
                // $(this).removeClass('text-theme text-theme2').addClass(color)
                $(this).attr('data-amplified', updateValue);
                statusMessageIndex = statusMessageIndex == 2 ? 0 : statusMessageIndex+1;
                $(this).attr('title', `Click to ${statusMessage[statusMessageIndex]} this entry.`)
                entryItem.Amplified = updateValue;
                let labelId = $('input[name="entryFilter"]:checked').val();
                $('.entryFilterLabel[value="'+labelId+'"]').click();
            }).catch((error) => {
                console.error("Error updating entry:", error);
                showModal("Error", "An error occurred while updating the entry. Please try again.");
            })
        })
    })
    .on('change','#sorter',function(){
        let sort = $(this).val();
        let month = $('#monthFilter').val();
        processListItems(spItems,month,sort);
    })
    .on('click','.entry-generate',function(){
        let attrid = $(this).attr('data-id');
        console.log("Generating presentation for entry ID:", attrid);
        createPresentation(spItems[attrid]);
    })
    .on('click','#installBtn', async function () {
        if (!deferredPrompt) return;

        deferredPrompt.prompt();    // Show the install prompt
        const choiceResult = await deferredPrompt.userChoice;
        console.log('User response:', choiceResult.outcome);

        deferredPrompt = null;      // Clear the stored prompt

    })
});