function showModal(title, message){
    $('#modal-title').text(title);
    $('#modal-message').html(message);
    $('#generalModal').modal('show');
}
function showConfirmModal(message,func){
    $('#confirmMessage').html(message);
    // Remove any previous click event bindings from the button
    $('#confirmModalBtn').off('click');

    // Bind the provided function to the button's click event
    $('#confirmModalBtn').on('click', func);

    // Show the modal
    $('#confirmModal').modal('show');
}
function addFormValidation(){
    'use strict'

    // Fetch all the forms we want to apply custom Bootstrap validation styles to
    const forms = document.querySelectorAll('.needs-validation')

    // Loop over them and prevent submission
    Array.from(forms).forEach(form => {
        form.addEventListener('submit', event => {
            if (!form.checkValidity()) {
                event.preventDefault()
                event.stopPropagation()
            }

            form.classList.add('was-validated')
        }, false)
    })
}
function resetFormValidation(form){
    form.removeClass('was-validated');
}
function spinButton(button){
    button.html('<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Loading...');
    button.prop('disabled', true);
}
function stopSpinButton(button, text){
    button.html(text);
    button.prop('disabled', false);
} 
function loadingHTML(){
    return `<div class="w-100 h-100 d-flex align-items-center justify-content-center">
        <div class="spinner-grow spinner-grow-sm text-primary mx-1" role="status"></div>
        <div class="spinner-grow spinner-grow-sm text-primary mx-1" role="status"></div>
        <div class="spinner-grow spinner-grow-sm text-primary mx-1" role="status"></div>
    </div>`;
}
function noContentHTML(){
    return `<div class="w-100 h-100 d-flex align-items-center justify-content-center">
        There is nothing here.
    </div>`;
}

function convertFileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result); // Get Base64 string
        reader.onerror = reject;
        reader.readAsDataURL(file); // Convert file to Base64
    });
}
function getRandomString(length) {
    const characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    let result = "";
    for (let i = 0; i < length; i++) {
        result += characters.charAt(Math.floor(Math.random() * characters.length));
    }
    return result;
}

function dummydata() {
    $('#entryInput').val("Dummy Entry");
    $('#subslDropdown').val("AMS-EU");
    $('#accountDropdown').val("Airbus");
    $('#teamInput').val("Dummy Team");
    $('#typeOfWork').val("Backup");
    let pill = `<div class="recipientPill bg-secondary-subtle px-2" value="mohamed.awaad@dxc.com" name="Awaad, Mohamed (DXC Luxoft)">Awaad, Mohamed (DXC Luxoft)<button class="removeRecipient" type="button"><i class="fa-solid fa-xmark"></i></button></div>`;
    $('#recipientContainer').html(pill);
    $('#recognitionInput1').attr('checked', true).prop('checked', true);
    $('#businessChallengeInput').val("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin ornare pharetra orci ac molestie. Vivamus at consectetur tortor. Nunc mattis orci non arcu fermentum, ac cursus mauris consequat. Etiam ligula nisi, vestibulum et viverra at, sagittis vel urna. Pellentesque a congue mauris. Aenean sed rutrum felis. Aliquam sit amet diam ac orci accumsan elementum. Vivamus imperdiet facilisis bibendum. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos.")
    $('#howDXCHelpedInput').val("In hac habitasse platea dictumst. Integer congue egestas nibh, non porta ex finibus eu. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Proin non varius risus. Aenean consequat nisl neque, vel convallis nisl tempor ac. Donec finibus tortor quis augue euismod interdum. Maecenas in lorem sit amet massa semper pulvinar. In quis rutrum enim. Morbi quis rutrum magna. Sed sodales ultrices pellentesque. Integer orci tortor, tempus vitae dapibus vitae, ultricies vitae leo.")
    $('#businessImpactInput').val("Nunc hendrerit ornare sem, sit amet commodo nibh. Donec non arcu rhoncus, porttitor mi eu, feugiat est. Fusce vestibulum ante vitae interdum semper. Sed ipsum purus, posuere et urna et, ornare semper felis. Sed est dolor, facilisis at nunc nec, molestie viverra arcu. Aliquam velit libero, aliquam in magna quis, vulputate faucibus mi. Nunc vel ipsum vitae urna placerat rhoncus ac at justo. Sed tristique posuere erat eu malesuada.");
}

//POWER AUTOMATE FUNCTIONS
function createModifyEmail(beforeData, afterData){
    // Create table structure
    htmlTable = `
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <thead>
                <tr>
                    <th style="padding: 8px; text-align: left;">Field</th>
                    <th style="padding: 8px; text-align: left;">Before</th>
                    <th style="padding: 8px; text-align: left;">After</th>
                    <th style="padding: 8px; text-align: left;">Modified By</th>
                    <th style="padding: 8px; text-align: left;">Modified On</th>
                </tr>
            </thead>
            <tbody>
    `;

    // Loop through modified fields
    afterData.forEach(({ name, value }) => {
        let formattedName = name;
        if(name == "Recipient_x0020_Emails"){
            formattedName = "Recipient Emails";
        }else if(name == "Recipient_x0020_Names"){
            formattedName = "Recipient Names";
        }
        htmlTable += `
            <tr>
                <td style="padding: 8px;">${formattedName}</td>
                <td style="padding: 8px;">${beforeData[name] || "N/A"}</td>
                <td style="padding: 8px;">${value}</td>
                <td style="padding: 8px;">${accountName}</td>
                <td style="padding: 8px;">${new Date().toLocaleString()}</td>
            </tr>
        `;
    });

    // Close table
    htmlTable += `
            </tbody>
        </table>
    `;
    sendEmail(htmlTable,"modified");
}
function createDeleteEmail(entryTitle){
    var html = `Entry <b>${entryTitle}</b> has been deleted by <b>${accountName}</b> on  ${new Date().toLocaleString()}.`;
    sendEmail(html,"modified");

}
async function sendEmail(email, action) {
    
    const requestData = {email: email, action: action};
    // const token = await getToken();
    flowUrl = "https://prod-59.westus.logic.azure.com:443/workflows/cd16bc1b918a4169b7b97fbe081aeb51/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=lsQ_obP9XwOg6WZ_-2FqJK3jw6sAX0ZbRLQHa7aQRhA"
    fetch(flowUrl, {
        method: "POST",
        headers: { 
            // "AUthorization": `Bearer ${token}`,
            "Content-Type": "application/json" 
        },
        body: JSON.stringify(requestData)

    })
    .then(async response => {
        console.log("Status Code:", response.status);

        if (!response.ok) {
            console.error("Error response from Power Automate:", response.statusText);
            return;
        }

        // Check if there's a response body
        const contentType = response.headers.get("Content-Type");
        
        if (contentType && contentType.includes("application/json")) {
            const data = await response.json();
            console.log("JSON Response:", data);
        } else {
            console.log("No JSON body, raw response:", await response.text());
        }
    })
    .catch(error => console.error("Error:", error));
}

function cleanPPTText(text){
    const cleanedText = text
    .replace(/[\x00-\x1F\x7F]/g, '') // removes most control characters
    .trim();
    return cleanedText;
}
function getRandomIndex(array) {
    return Math.floor(Math.random() * array.length);  
}

function callTippy(selector,content="Fancy seeing you here!",placement="top"){
    const el = $(selector)[0];

    if (el._tippy) {
        el._tippy.setContent(content);
        el._tippy.setProps({
            placement, placement,
            arrow: true,
            allowHTML: true,
            interactive: true,
            trigger: 'manual'
        })
        el._tippy.show();
    } else {        
        tippy(selector, {
            content: content,
            placement, placement,
            arrow: true,
            allowHTML: true,
            interactive: true,
            trigger: 'manual'
        }).show();
    }
}

$(document).ready(function(){
    $(document)
    .on('click','#confirmModalCancel',function(){
        $('#confirmModalBtn').off('click');
    })
    .on('click','.help-btn',function(){
        $('#helpCanvas .offcanvas-title').text($(this).data('title'));
        let carousel = $(this).attr('data-carousel');
        $('#helpCanvas .carousel').hide();
        $('#'+carousel).show();
        let prev = $(this).attr('data-previous');
        $('#helpCanvas .btn-close').attr("data-bs-target","#"+prev).attr("aria-controls",prev);
        if(access.type == 2){
            $('#entryInfoCarousel .carousel-item').removeClass('active');
            $('#entryInfoCarousel .carousel-item').first().addClass('active');
        }
    })
    .on('click', '#cycleBG', function() {
        let currentBG = parseInt($('#cycleBG').attr('backgroundImage'));
        currentBG = currentBG === 4 ? 1 : currentBG + 1; // Cycle through backgrounds 1 to 4
        
        $('#home-page').css('backgroundImage', 'url("assets/img/bg' + currentBG + '.png")');
        $(this).css('transform', 'rotate(' + (currentBG * 90) + 'deg)');
        $(this).attr('backgroundImage', currentBG);
    })
    .on('click','.info-button',function() {
        const infoType = $(this).attr('id').replace('Info', '');
        let message = '';
        let title='';
        switch (infoType) {
            case 'recognition':
                message = `
                    <p>Individual / Team recognition is determined based on the number of recipients.</p>`;
                title = 'Individual / Team Recognition';
                break;
            case 'recipientNames':
                message = `
                    <p>Search the recipients on the search bar. If the recipient has already resigned or can't be found, click on <b>"Others"</b>. <br>Similarly, you can also add group or team names using <b>"Others"</b>.</p>
                    <p><b>Note:</b> Adding or changing recipients will automatically reset uploaded images.</p>
                    `;
                title = 'Recipient Names';
                break;
            case 'recipientImages':
                message = `
                    <p>Upload images of the recipients. You can select multiple images.<br></p>
                    <p><b>Note:</b> You can only upload one image per recipient, upto a maximum of 4 images. </p>
                    <p>For more than 4 recipients, it is recommended to upload <b>one</b> image in landscape orientation.</p>
                    `;
                title = 'Recipient Images';
                break;
            case 'businessChallenge':
                message = `
                Should answer the question “What happened”?
                <ol>
                            <li>What is the background of the issue? Describe the situation</li>
                            <li>What issue needed to be solved/addressed? What problem caused the need for DXC to take action?</li>
                            <li>What is the timeline or timeframe? Did something need to be resolved quickly? Was there an issue going on for too long?</li>
                        </ol>`;
                        title = 'Customer Challenge';
                        break;
                    case 'howDXCHelped':
                        message = 
                        `Should answer the question “What did DXC do to execute the task and how we did it?”
                        <br>What specific actions were taken?`;
                        title = 'How DXC Helped';
                        break;
                    case 'businessImpact':
                        message = `
                            Should answer “What was the impact to the customer?” or “How does it advance or push forward our customer first strategy/ initiative?”<br>
                            What would have happened on the customer's end if you were not successful or the action by DXC was not done?<br>
                            Did it support any of your customer's initiatives?<br>
                            How did this help the customer?<br>
                            Are there any cost savings to the customer?<br>
                            Did it save time and efforts to the customer?<br>
                            Did it give way to an opportunity for renewal of work/ contract extension or for additional work?<br>
                            <br>
                            Quantify Results
                            <ul>
                                <li>Cost savings of X amount/quarter to the customer</li>
                                <li>X amount of revenue loss of the customer if the issue was not resolved by DXC</li>
                                <li>Reduced effort from X hours to Y minutes</li>
                                <li>Reduced tickets by x % in Y months</li>
                            </ul>`;
                        title = 'Customer Impact';
                        break;
                    case 'uniqueness':
                        message = 
                        `Should answer the question “What sets this apart from others?”
                            <br>Is the action/task unique? What makes this unique? Has this not been done before?
                            <br>What uniqueness does it introduce to the Customer/ Market?
                            <br>What makes this complex? Is it very technical? Is it due to time constraints? Is it due to volume of work? Are there multiple factors that need to be considered simultaneously? 
                            <br>Is this an innovation?
                            <br>NON-BAU is preferred
                            <br>Can be BAU work as long as uniqueness or complexity is justified
                            <br>Client centric –impact statement
                            <br>The amazing stories that have been justified as unique or complex will be reviewed with Cleif and can potentially become candidates for submission to global.`
                        ;
                        title = 'Uniqueness/Complexity';
                        break;
                    case 'outcome':
                        message = 'Describe the business outcome achieved through this work.';
                        title = 'Business Outcome';
                        break;
                    default:
                        message = 'No information available.';
                }
                showModal(title, message);
    })

            
})