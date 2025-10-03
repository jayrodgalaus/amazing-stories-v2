//ppt variables
var themeColor = "5F249F";
var titleFont = "Arial";

class ChallengeImage{
    constructor(x, y) {
        this.path= "assets/img/ppt/business-challenge.png";
        this.x = cmToIn(x);  // X position in inches
        this.y = cmToIn(y);  // Y position in inches
        this.w = cmToIn(1.41);  // Width in inches
        this.h = cmToIn(1.35);  // Height in inches
    }
}
class ImpactImage{
    constructor(x, y) {
        this.path= "assets/img/ppt/business-impact.png";
        this.x = cmToIn(x);  // X position in inches
        this.y = cmToIn(y);  // Y position in inches
        this.w = cmToIn(1.21);  // Width in inches
        this.h = cmToIn(1.38);  // Height in inches
    }
}
class HelpImage{
    constructor(x, y) {
        this.path= "assets/img/ppt/help.png";
        this.x = cmToIn(x);  // X position in inches
        this.y = cmToIn(y);  // Y position in inches
        this.w = cmToIn(1.29);  // Width in inches
        this.h = cmToIn(1.35);  // Height in inches
    }
}
class Divider{
    constructor(y) {
        this.x = cmToIn(12.5);
        this.y = cmToIn(y);
        this.line = {color : "BFBFBF"}
        this.w = cmToIn(20.7);
        this.h = cmToIn(0.01); // Height of the line
    }
}

async function createPresentation(entry){
    console.log(entry);
    showModal("Generating Presentation", "Please wait while the presentation is being generated. Your download should start automatically. <br><br><b>Note:</b> If there are missing images or text, or if the download doesn't start, please refresh the page and try again.");
    const {template, pages} = getTemplateType(entry);
    const entryId = entry.Id;
    let recipientCount = entry.Recipients ? entry.Recipients.split('/ ').length : 0;
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.theme = { bodyFontFace: "Arial" };
    let slides = [];
    const bgImage = {
        path: "assets/img/ppt/background-graphic.png",
        x: cmToIn(1.6),  // Horizontal position (in inches)
        y: cmToIn(4.76),  // Vertical position (in inches)
        w: cmToIn(9.77),  // Width (in inches)
        h: cmToIn(8.98)   // Height (in inches)
    }
    const dxclabel = {
        path: "assets/img/ppt/dxclabel.png",
        x: cmToIn(1.6),  // Horizontal position (in inches)
        y: cmToIn(17.54),  // Vertical position (in inches)
        w: cmToIn(5.26),  // Width (in inches)
        h: cmToIn(0.62)   // Height (in inches)
    }
    let logoImage = {
        path: "",
        x : cmToIn(28.38),  // X position in inches
        y : cmToIn(0.77),  // Y position in inches
        sizing: { 
            type: 'contain', 
            w: cmToIn(4.8), 
            h: cmToIn(2.5)
        } // Adjusted to fit the logo
    }
    let challengeImage = new ChallengeImage(12.76, 4.85);
    let impactImage = new ImpactImage(13.03, 9.04);
    let uniquenessImage = new ImpactImage(13.03, 9.04); // Using impactImage for Uniqueness
    let helpImage = new HelpImage(12.93, 13.24);
    var accountBox = {
        x: cmToIn(1.59),  // X position in inches
        y: cmToIn(13.75),  // Y position in inches
        w: cmToIn(9.77),  // Width in inches
        h: cmToIn(3.48),  // Height in inches
        fill: { color: "D9D9D6" }, // Example red fill color
        isTextBox: true,
        margin: {
            top: cmToIn(0.64), // Top margin in inches
            right: cmToIn(0.64), // Right margin in inches
            bottom: cmToIn(0.64), // Bottom margin in inches
            left: cmToIn(0.64) // Left margin in inches
        },
    }
    var subslbox = {
        x: cmToIn(1.59),  // X position in inches
        y: cmToIn(1.13),  // Y position in inches
        w: cmToIn(5.54),  // Width in inches
        h: cmToIn(1.45),  // Height in inches
        isTextBox: true,
        align: "left",
        fit: "shrink",
        fontSize: 28,
        color: themeColor,
        bold: true,
    }
    var titlebox = {
        x: cmToIn(7.12),  // X position in inches
        y: cmToIn(1.13),  // Y position in inches
        w: cmToIn(21.08),  // Width in inches
        h: cmToIn(2.65),  // Height in inches
        isTextBox: true,
        align: "center",
        fit: "shrink",
        fontSize: 28,
        color: themeColor,
        bold: true,
        fontFace: titleFont,
        valign: "top"
    }
    var indivRecipientBox = { // default individual recipient box
        x: cmToIn(1.59),  // X position in inches
        y: cmToIn(7.1),  // Y position in inches
        w: cmToIn(9.77),  // Width in inches
        h: cmToIn(0.6),  // Height in inches
        isTextBox: true,
        align: "center",
        fit: "shrink",
        fontSize: 14,
        color: 'FFFFFF'
    }
    var team4RecipientBox = { // default team recipient box
        x: cmToIn(1.59),  // X position in inches
        y: cmToIn(5),  // Y position in inches
        w: cmToIn(9.75),  // Width in inches
        h: cmToIn(2),  // Height in inches
        isTextBox: true,
        align: "center",
        fit: "shrink",
        fontSize: 12,
        color: 'FFFFFF',
        valign: "bottom"
    }
    var team10RecipientBox = { // default team recipient box
        x: cmToIn(1.59),  // X position in inches
        y: cmToIn(5),  // Y position in inches
        w: cmToIn(9.75),  // Width in inches
        h: cmToIn(4.45),  // Height in inches
        isTextBox: true,
        align: "center",
        fontSize: 10,
        color: 'FFFFFF',
        valign: "bottom"
    }
    var team17RecipientBox = { // default team recipient box
        x: cmToIn(1.59),  // X position in inches
        y: cmToIn(5),  // Y position in inches
        w: cmToIn(9.75),  // Width in inches
        h: cmToIn(8.75),  // Height in inches
        isTextBox: true,
        align: "center",
        fit: "shrink",
        fontSize: 11,
        color: 'FFFFFF',
        valign: "center"
    }
    var teamHalfBox1={
        x: cmToIn(1.59),  // X position in inches
        y: cmToIn(5),  // Y position in inches
        w: cmToIn(4.875),  // Width in inches
        h: cmToIn(8.75),  // Height in inches
        isTextBox: true,
        align: "left",
        fit: "shrink",
        fontSize: 11,
        color: 'FFFFFF',
        valign: "top"
    }
    var teamHalfBox2={
        x: cmToIn(6.47),  // X position in inches
        y: cmToIn(5),  // Y position in inches
        w: cmToIn(4.875),  // Width in inches
        h: cmToIn(8.75),  // Height in inches
        isTextBox: true,
        align: "left",
        fit: "shrink",
        fontSize: 11,
        color: 'FFFFFF',
        valign: "top"
    }
    var teamHalfBox11={
        x: cmToIn(1.59),  // X position in inches
        y: cmToIn(4.8),  // Y position in inches
        w: cmToIn(4.875),  // Width in inches
        h: cmToIn(5.25),  // Height in inches
        isTextBox: true,
        align: "left",
        fit: "shrink",
        fontSize: 10,
        color: 'FFFFFF',
        valign: "top"
    }
    var teamHalfBox24={
        x: cmToIn(6.47),  // X position in inches
        y: cmToIn(4.8),  // Y position in inches
        w: cmToIn(4.875),  // Width in inches
        h: cmToIn(5.25),  // Height in inches
        isTextBox: true,
        align: "left",
        fit: "shrink",
        fontSize: 10,
        color: 'FFFFFF',
        valign: "top"
    }
    var challengeBox = {
        x: cmToIn(14.3),  // X position in inches
        y: cmToIn(4.64),  // Y position in inches
        w: cmToIn(18.6),  // Width in inches
        h: cmToIn(4.05),  // Height in inches
        isTextBox: true,
        align: "left",
        valign: "top"
    }
    var impactBox = {
        x: cmToIn(14.3),  // X position in inches
        y: cmToIn(8.89),  // Y position in inches
        w: cmToIn(18.6),  // Width in inches
        h: cmToIn(4.05),  // Height in inches
        isTextBox: true,
        align: "left",
        valign: "top"
    }
    var helpBox = {
        x: cmToIn(14.3),  // X position in inches
        y: cmToIn(13.12),  // Y position in inches
        w: cmToIn(18.6),  // Width in inches
        h: cmToIn(4.05),  // Height in inches
        isTextBox: true,
        align: "left",
        valign: "top"
    }
    var uniquenessBox = {
        x: cmToIn(14.3),  // X position in inches
        y: cmToIn(8.89),  // Y position in inches
        w: cmToIn(9.3),  // Width in inches
        h: cmToIn(4.05),  // Height in inches
        isTextBox: true,
        align: "left",
        valign: "top"
    }
    var copyrightBox = {
        x: cmToIn(12.71),  // X position in inches
        y: cmToIn(17.61),  // Y position in inches
        w: cmToIn(8.46),  // Width in inches
        h: cmToIn(0.61),  // Height in inches
        isTextBox: true,
        align: "center",
        valign: "top",
    }
    var internalBox = {
        x: cmToIn(16.01),  // X position in inches
        y: cmToIn(18.25),  // Y position in inches
        w: cmToIn(1.89),  // Width in inches
        h: cmToIn(0.51),  // Height in inches
        isTextBox: true,
        align: "center",
        valign: "top",
    }
    let images = await getPptRecipientImages(entryId);
    let accountText = [
        { text: "Group Name / Account Name / Team Name:\n", options: { bold: false, fontSize: 11 } },
        { text: `${entry.SUBSL} / ${entry.Team}\n\n`, options: { bold: true, fontSize: 11 } },
        { text: "Type of Work: ", options: { bold: false, fontSize: 11 } },
        { text: `${entry.Worktype}\n\n`, options: { bold: true, fontSize: 11 } },
        { text: "Month: ", options: { bold: false, fontSize: 11 } },
        { text: `${entry.Month}`, options: { bold: true, fontSize: 11 } }
    ];
    let subslText = [
        { text: `[${entry.SUBSL}]`, options: {fit: "shrink" } }
    ];
    let challengeText = [
        { text: "Customer Challenge\n", options:{ bold: true, fontSize: 16, color: themeColor}},
        { text: cleanPPTText(entry.Challenge), options: { fontSize: 12, color: '000000' } }
    ];
    let impactText = [
        { text: "Customer Impact\n", options:{ bold: true, fontSize: 16, color: themeColor}},
        { text: cleanPPTText(entry.Impact), options: { fontSize: 12, color: '000000' } }
    ];
    let uniquenessText = [];
    if(entry.Uniqueness && entry.Uniqueness.trim().length > 0){
        uniquenessText = [
            { text: "Uniqueness/Complexity\n", options:{ bold: true, fontSize: 16, color: themeColor}},
            { text: cleanPPTText(entry.Uniqueness), options: { fontSize: 12, color: '000000' } }
        ];
    }
    let helpText = [
        { text: "How DXC Helped\n", options:{ bold: true, fontSize: 16, color: themeColor}},
        { text: cleanPPTText(entry.Help), options: { fontSize: 12, color: '000000' } }
    ];
    let copyrightText = [
        { text: "© 2023 DXC Technology Company. All rights reserved.", options: { fontSize: 9.2, color: '000000' } }
    ]
    let internalText = [
        { text: "DXC Internal", options: { fontSize: 6.8, color: '000000' } }
    ]
    
    //split recipients into array text
    let recipientsArray = splitRecipientText(entry.Recipients);
    console.log(template, pages, recipientCount);
    //adding duplicate info per slide
    for(i =0; i<pages; i++){
        slides[i] = pptx.addSlide();
        slides[i].addImage(bgImage);
        slides[i].addImage(dxclabel);
        slides[i].addText(accountText, { ...accountBox });
        slides[i].addText(subslText, { ...subslbox });
        slides[i].addText(entry.Title, { ...titlebox });
        slides[i].addText(copyrightText, { ...copyrightBox });
        slides[i].addText(internalText, { ...internalBox });
        let logopath = getAccountLogo(entry.Account);
        if(logopath){
            logoImage.path = logopath;
            let logodimensions = await getImageDimensions(logopath);
            let logowidth = pxToInches(logodimensions.width);
            let logoheight = pxToInches(logodimensions.height);
            logoImage.w = logowidth;
            logoImage.h = logoheight;
            slides[i].addImage(logoImage);
        }
        // adding recipient images per slide
        addRecipientImages(slides[i], images, recipientCount);
        // adding recipient text per slide
        if(recipientCount > 0){
            if(recipientCount === 1){
                slides[i].addText(recipientsArray, { ...indivRecipientBox });
            }else if (recipientCount > 1 && recipientCount <= 4) {
                if(recipientCount == 4 && images.length < 4){ // if only one recipient but 4 recipients in the field
                    if(images.length == 1 || images.length == 2)
                        slides[i].addText(recipientsArray, { ...team4RecipientBox });
                    else
                        slides[i].addText(recipientsArray, { ...indivRecipientBox });
                }else{
                    slides[i].addText(recipientsArray, { ...team4RecipientBox });
                }
            }else if (recipientCount > 4 && recipientCount <= 10 && images.length == 1){
                slides[i].addText(recipientsArray, { ...team10RecipientBox });
            }else if (recipientCount > 10 && recipientCount <= 24 && images.length == 1){
                let mid = Math.ceil(recipientsArray.length / 2); // Round up to ensure even splits
                let firstHalf = recipientsArray.slice(0, mid);
                let secondHalf = recipientsArray.slice(mid);
                slides[i].addText(firstHalf, { ...teamHalfBox11 });
                slides[i].addText(secondHalf, { ...teamHalfBox24 });
            }else if (recipientCount > 4 && recipientCount <= 17 && images.length == 0){
                slides[i].addText(recipientsArray, { ...team17RecipientBox });
            }else if( recipientCount > 17 && recipientCount <= 34  && images.length == 0){
                let mid = Math.ceil(recipientsArray.length / 2); // Round up to ensure even splits
                let firstHalf = recipientsArray.slice(0, mid);
                let secondHalf = recipientsArray.slice(mid);
                slides[i].addText(firstHalf, { ...teamHalfBox1 });
                slides[i].addText(secondHalf, { ...teamHalfBox2 });
            }else if( recipientCount > 34 && recipientCount <= 40 && images.length == 0){
                let mid = Math.ceil(recipientsArray.length / 2); // Round up to ensure even splits
                let firstHalf = recipientsArray.slice(0, mid);
                let secondHalf = recipientsArray.slice(mid);
                teamHalfBox1.fontSize = 9.5;
                teamHalfBox2.fontSize = 9.5;
                slides[i].addText(firstHalf, { ...teamHalfBox1 });
                slides[i].addText(secondHalf, { ...teamHalfBox2 });
            }else{
                let error = `No template set for ${recipientCount} width ${images.length} images.`;
                showModal("Error", error);
                return;
            }
        }
    }
    switch(template){
        case "sp":{ //single page without uniqueness
            slides[0].addImage(challengeImage);
            slides[0].addImage(impactImage);
            slides[0].addImage(helpImage);
            slides[0].addText(challengeText, { ...challengeBox });
            slides[0].addText(impactText, { ...impactBox });
            slides[0].addText(helpText, { ...helpBox });
            let div1 = new Divider(8.8);
            let div2 = new Divider(13);
            slides[0].addShape(pptx.shapes.LINE, div1);
            slides[0].addShape(pptx.shapes.LINE, div2);
            break;}
        case "spu":{ //single page with uniqueness
            slides[0].addImage(challengeImage);
            slides[0].addImage(impactImage);
            slides[0].addImage(helpImage);
            slides[0].addText(challengeText, { ...challengeBox });
            impactBox.w = cmToIn(9.3);
            slides[0].addText(impactText, { ...impactBox });
            if(entry.Uniqueness && entry.Uniqueness.trim().length > 0){
                uniquenessBox.x = cmToIn(23.7);
                uniquenessBox.w = cmToIn(9.3);
                slides[0].addText(uniquenessText, { ...uniquenessBox });
                slides[0].addImage(uniquenessImage);
            }
            slides[0].addText(helpText, { ...helpBox });
            let div1 = new Divider(8.8);
            let div2 = new Divider(13);
            slides[0].addShape(pptx.shapes.LINE, div1);
            slides[0].addShape(pptx.shapes.LINE, div2);
            break;}
        case "dp":{ //double page without uniqueness
            //page 1
            slides[0].addImage(challengeImage);
            impactImage.x = cmToIn(13.03);
            impactImage.y = cmToIn(11.52);
            slides[0].addImage(impactImage);
            slides[0].addText(challengeText, { ...challengeBox });
            impactBox.y = cmToIn(11.38);
            slides[0].addText(impactText, { ...impactBox });
            let div1 = new Divider(11.24);
            slides[0].addShape(pptx.shapes.LINE, div1);
            //page 2
            helpImage.x = cmToIn(13.01);
            helpImage.y = cmToIn(4.88);
            slides[1].addImage(helpImage);
            helpBox.y = cmToIn(4.64);
            helpBox.h = cmToIn(12.6);
            slides[1].addText(helpText, { ...helpBox });
            break;}
        case "dpu_short":{ //double page with uniqueness (short)
            //page 1
            slides[0].addImage(challengeImage);
            impactImage.y = cmToIn(11.52); // Adjusting position for double page
            slides[0].addImage(impactImage);
            slides[0].addText(challengeText, { ...challengeBox });
            impactBox.y = cmToIn(11.38);
            impactBox.w = cmToIn(9.3);
            slides[0].addText(impactText, { ...impactBox });
            let div1 = new Divider(11.24);
            slides[0].addShape(pptx.shapes.LINE, div1);
            if(entry.Uniqueness && entry.Uniqueness.trim().length > 0){
                uniquenessBox.x = cmToIn(23.6);
                uniquenessBox.y = cmToIn(11.38);
                uniquenessBox.w = cmToIn(9.3);
                slides[0].addText(uniquenessText, { ...uniquenessBox });
            }
            //page 2
            helpImage.y = cmToIn(4.88); // Adjusting position for double page
            slides[1].addImage(helpImage);
            helpBox.y = cmToIn(4.64);
            helpBox.h = cmToIn(12.6);
            slides[1].addText(helpText, { ...helpBox });
            break;}
        case "dpu_long":{ //double page with uniqueness (long)
            //page 1
            slides[0].addImage(challengeImage);
            impactImage.y = cmToIn(11.52); // Adjusting position for double page
            slides[0].addImage(impactImage);
            helpImage.y = cmToIn(4.88); // Adjusting position for double page
            slides[0].addText(challengeText, { ...challengeBox });
            impactBox.y = cmToIn(11.38);
            impactBox.w = cmToIn(18.6);
            slides[0].addText(impactText, { ...impactBox });
            let div1 = new Divider(11.24);
            slides[0].addShape(pptx.shapes.LINE, div1);
            //page 2
            if(entry.Uniqueness && entry.Uniqueness.trim().length > 0){
                uniquenessImage.x = cmToIn(13.03);
                uniquenessImage.y = cmToIn(4.91);
                slides[1].addImage(uniquenessImage);
                uniquenessBox.x = cmToIn(14.3);
                uniquenessBox.y = cmToIn(4.64);
                uniquenessBox.w = cmToIn(18.6);
                slides[1].addText(uniquenessText, { ...uniquenessBox });
            }
            helpImage.x = cmToIn(13.01);
            helpImage.y = cmToIn(11.52);
            slides[1].addImage(helpImage);
            helpBox.y = cmToIn(11.38);
            slides[1].addText(helpText, { ...helpBox });
            let div2 = new Divider(11.24);
            slides[1].addShape(pptx.shapes.LINE, div2);
            break;}
        case "tp":{ //triple page
            let topy = cmToIn(4.64);
            let fullwidth = cmToIn(18.6);
            //page 1
            slides[0].addImage(challengeImage);
            slides[0].addText(challengeText, { ...challengeBox });
            //page 2
            impactImage.x = cmToIn(13.03);
            impactImage.y = cmToIn(4.91);
            slides[1].addImage(impactImage);
            impactBox.y = topy;
            impactBox.w = fullwidth;
            slides[1].addText(impactText, { ...impactBox });
            //page 3
            helpImage.x = cmToIn(13.01);
            helpImage.y = cmToIn(4.88);
            slides[2].addImage(helpImage);
            helpBox.y = topy;
            helpBox.w = fullwidth;
            slides[2].addText(helpText, { ...helpBox });

            break;}
        case "full":{ //full page
            let topy = cmToIn(4.64);
            let fullwidth = cmToIn(18.6);
            //page 1
            slides[0].addImage(challengeImage);
            slides[0].addText(challengeText, { ...challengeBox });
            //page 2
            impactImage.x = cmToIn(13.03);
            impactImage.y = cmToIn(4.91);
            slides[1].addImage(impactImage);
            impactBox.y = topy;
            impactBox.w = fullwidth;
            slides[1].addText(impactText, { ...impactBox });
            //page 3
            if(entry.Uniqueness && entry.Uniqueness.trim().length > 0){
                uniquenessImage.x = cmToIn(13.03);
                uniquenessImage.y = cmToIn(4.91);
                slides[2].addImage(uniquenessImage);
                uniquenessBox.x = cmToIn(14.3);
                uniquenessBox.y = topy;
                uniquenessBox.w = fullwidth;
                slides[2].addText(uniquenessText, { ...uniquenessBox });
            }
            //page 4
            helpImage.x = cmToIn(13.01);
            helpImage.y = cmToIn(4.88);
            slides[3].addImage(helpImage);
            helpBox.y = topy;
            helpBox.w = fullwidth;
            slides[3].addText(helpText, { ...helpBox });
            break;}
    }
    pptx.writeFile({ fileName: "BrowserPresentation.pptx" });
}
function adjustedName(name) {
    if (name.length <= 22) return name;

    if (name.includes(",")) {
        // Format: Last Name, First Name
        let [lastName, firstName] = name.split(",");
        let firstWord = firstName.trim().split(" ")[0]; // Get first word only
        return `${lastName.trim()}, ${firstWord}`;
    } else {
        // Format: First Name Last Name
        let words = name.trim().split(" ");
        let first = words[0];
        let last = words[words.length - 1];
        return `${first} ${last}`;
    }
}

function cmToIn(cm) {
    return cm * 0.3937; // Convert cm to inches
}
function pxToCm(px) {
    return px * (2.54 / 96); // Assuming 96 PPI
}
function pxToInches(px, ppi = 96) {
    return px / ppi;
}

function splitRecipientText(recipients) {
    return recipients.split("/").map((name, index, arr) => ({
        text: index === arr.length - 1 ? adjustedName(name).trim() : adjustedName(name).trim() + "\n"
    }));
}


function getImageDimensionsFromBase64(base64String) {
    return new Promise((resolve) => {
        let img = new Image();
        img.onload = () => resolve({ width: img.width, height: img.height });
        img.src = base64String;
    });
}

function getImageDimensionsFromURL(imageUrl) {
    return new Promise((resolve, reject) => {
        let img = new Image();
        img.onload = () => resolve({ width: img.width, height: img.height });
        img.onerror = () => reject("❌ Error loading image.");
        img.src = imageUrl;
    });
}
function getImageDimensions(source){
    return new Promise((resolve, reject) => {
        let img = new Image();
        img.onload = () => resolve({ width: img.width, height: img.height });
        img.onerror = () => reject("❌ Error loading image.");
        img.src = source;
    });
}
function addRecipientImages_old(slide, images, recipientCount) {
    let sizingOptions = { 
        type: 'contain',
        w: cmToIn(4),//,
        h: cmToIn(4)//
    }
    if (recipientCount === 1) {
        if (Array.isArray(images) && images.length > 0) {
            let dimensions = getImageDimensionsFromBase64(images[0]);
            let imgwidth = pxToInches(dimensions.width);
            let imgheight = pxToInches(dimensions.height);
            slide.addImage({
                data: images[0],
                x: cmToIn(4.485),
                y: cmToIn(8),
                w: imgwidth,
                h: imgheight,
                sizing: sizingOptions
            });
        }
    } else if (recipientCount === 2) {
        if (Array.isArray(images) && images.length > 0) {
           if( images.length === 1) { // two recipients but only one image
                let dimensions = getImageDimensionsFromBase64(images[0]);
                let imgwidth = pxToInches(dimensions.width);
                let imgheight = pxToInches(dimensions.height);
                slide.addImage({
                    data: images[0],
                    x: cmToIn(4.485),
                    y: cmToIn(8),
                    w: imgwidth,
                    h: imgheight,
                    sizing: sizingOptions
                });
            }else{
                for (let j = 0; j < images.length; j++) {
                    let imgOptions = {
                        data: images[j],
                        x: cmToIn(2.02), // Adjusting X position for each image
                        y: cmToIn(8),
                        sizing: sizingOptions
                    }
                    let dimensions = getImageDimensionsFromBase64(images[j]);
                    let imgwidth = pxToInches(dimensions.width);
                    let imgheight = pxToInches(dimensions.height);
                    imgOptions.w = imgwidth;
                    imgOptions.h = imgheight;
                    if(j === 1){//2nd image
                        imgOptions.x = cmToIn(6.89);
                        imgOptions.y = cmToIn(8);
                    }
                    slide.addImage(imgOptions);
                }
            }
        }
    } else if (recipientCount === 3){
        if (Array.isArray(images) && images.length > 0) {
           if( images.length === 1) { // two recipients but only one image
                let dimensions = getImageDimensionsFromBase64(images[0]);
                let imgwidth = pxToInches(dimensions.width);
                let imgheight = pxToInches(dimensions.height);
                //check if image is landscape or portrait
                
                slide.addImage({
                    data: images[0],
                    x: cmToIn(4.485),
                    y: cmToIn(8),
                    w: imgwidth,
                    h: imgheight,
                    sizing: sizingOptions
                });
            }else if (images.length === 2) {
                for (let j = 0; j < images.length; j++) {
                    let imgOptions = {
                        data: images[j],
                        x: cmToIn(2.02), // Adjusting X position for each image
                        y: cmToIn(8),
                        sizing: sizingOptions
                    }
                    let dimensions = getImageDimensionsFromBase64(images[j]);
                    let imgwidth = pxToInches(dimensions.width);
                    let imgheight = pxToInches(dimensions.height);
                    imgOptions.w = imgwidth;
                    imgOptions.h = imgheight;
                    if(j === 1){//2nd image
                        imgOptions.x = cmToIn(6.89);
                        imgOptions.y = cmToIn(8);
                    }
                    slide.addImage(imgOptions);
                }
            }else{
                for (let j = 0; j < images.length; j++) {
                    let imgOptions = {
                        data: images[j],
                        x: cmToIn(2.01), // Adjusting X position for each image
                        y: cmToIn(9),
                        sizing: sizingOptions
                    }
                    let dimensions = getImageDimensionsFromBase64(images[j]);
                    let imgwidth = pxToInches(dimensions.width);
                    let imgheight = pxToInches(dimensions.height);
                    imgOptions.w = imgwidth;
                    imgOptions.h = imgheight;
                    if(j === 1){//2nd image
                        imgOptions.x = cmToIn(5.07);
                        imgOptions.y = cmToIn(9);
                    }
                    else if(j === 2){//3rd image
                        imgOptions.x = cmToIn(8.14);
                        imgOptions.y = cmToIn(9);
                    }
                    slide.addImage(imgOptions);
                }
            }
        }
    }
}
async function addRecipientImages(slide, images, recipientCount) {
    let sizingOptions = getSizingOptions(recipientCount, images.length);

    if (!Array.isArray(images) || images.length === 0) return;

    if (images.length === 1) {
        if(recipientCount >= 1 && recipientCount <= 4){
            //check if image is landscape or portrait
            //let orientation = getImgOrientation(images[0].width, images[0].height);
            // If only one recipient or one image, place it in the center
            await addSingleImage(slide, images[0], cmToIn(4.485), cmToIn(8), sizingOptions);
        }else if(recipientCount >= 5 && recipientCount <= 10){
            //group photo
            await addSingleImage(slide, images[0], cmToIn(2.11), cmToIn(9.5), sizingOptions);
        }else if(recipientCount >= 11 && recipientCount <= 24){
            //team photo
            await addSingleImage(slide, images[0], cmToIn(2.64), cmToIn(10.14), sizingOptions);
        }
    } else {

        addMultipleImages(slide, images, recipientCount, sizingOptions);
    }
}

async function addSingleImage(slide, image, x, y, sizingOptions) {
    let dimensions = await getImageDimensionsFromBase64(image);
    slide.addImage({
        data: image,
        x: x,
        y: y,
        w: pxToInches(dimensions.width),
        h: pxToInches(dimensions.height),
        sizing: sizingOptions
    });
}

async function addMultipleImages(slide, images, recipientCount, sizingOptions) {
    const positions = getImagePositions(recipientCount, images.length);
    
    images.forEach(async (image, index) => {
        if (index >= positions.length) return; // Prevent excess images being placed

        let dimensions = await getImageDimensionsFromBase64(image);
        slide.addImage({
            data: image,
            x: positions[index].x,
            y: positions[index].y,
            w: pxToInches(dimensions.width),
            h: pxToInches(dimensions.height),
            sizing: sizingOptions
        });
        /* console.log(dimensions,{x: positions[index].x,
            y: positions[index].y,
            w: pxToInches(dimensions.width),
            h: pxToInches(dimensions.height),
            sizing: sizingOptions}) */
    });
}

function getImagePositions(recipientCount, imageslength) {
    console.log(`Recipient Count: ${recipientCount}, Images Length: ${imageslength}`);
    const positions = {
        2: [ { x: cmToIn(2.02), y: cmToIn(8) }, { x: cmToIn(6.89), y: cmToIn(8) } ],
        3: [ { x: cmToIn(2.01), y: cmToIn(9) }, { x: cmToIn(5.07), y: cmToIn(9) }, { x: cmToIn(8.14), y: cmToIn(9) } ],
        4: [ { x: cmToIn(3.27), y: cmToIn(7.17) }, { x: cmToIn(6.42), y: cmToIn(7.17) }, { x: cmToIn(3.27), y: cmToIn(10.3) }, { x: cmToIn(6.42), y: cmToIn(10.3) } ],
        group5_10: [{ x: cmToIn(2.11), y: cmToIn(9.5) }],
        group11_24: [{ x: cmToIn(2.64), y: cmToIn(10.14) }]
    };
    if(recipientCount <= 4 && imageslength == recipientCount) {
        return positions[recipientCount];
    }else if(recipientCount <= 4 && imageslength < recipientCount) {
        return positions[imageslength];
    }else if(recipientCount >= 5 && recipientCount <= 10){
        if(imageslength === 1) {
            return positions.group5_10;
        }else {
            console.warn(`⚠️ Recipient count ${recipientCount} contains ${imageslength} images.`);
            return [];
        }
    }else if(recipientCount >= 11 && recipientCount <= 24){
        if(imageslength === 1) {
            return positions.group11_24;
        }else {
            console.warn(`⚠️ Recipient count ${recipientCount} contains ${imageslength} images.`);
            return [];
        }
    }else if(recipientCount > 4 && imageslength > 1) {
        console.warn(`⚠️ Recipient count ${recipientCount} contains ${imageslength} images.`);
        return [];
    }

    return positions[recipientCount] || [];
}

function getSizingOptions(recipientCount, imageslength){
    let sizingOptions = { 
        type: 'cover',
        w: cmToIn(4),
        h: cmToIn(4)
    };

    if (recipientCount === 3 || recipientCount === 4) {
        sizingOptions.w = cmToIn(2.8);
        sizingOptions.h = cmToIn(2.8);
        if(imageslength === 1 || imageslength === 2) {
            sizingOptions.w = cmToIn(4);
            sizingOptions.h = cmToIn(4);
        }

    }else if(recipientCount >= 5 && recipientCount <= 10 && imageslength === 1){
        sizingOptions.type = 'contain';
        sizingOptions.w = cmToIn(8.77);
        sizingOptions.h = cmToIn(4);
    } else if(recipientCount >= 11 && recipientCount <= 24 && imageslength === 1) {
        sizingOptions.type = 'contain';
        sizingOptions.w = cmToIn(7.67);
        sizingOptions.h = cmToIn(3.5);

    }

    return sizingOptions;
}

function getImgOrientation(imgwidth, imgheight) {
    const threshold = 1.1; // Exclude nearly identical ratios (e.g., 1.01 would be too small a difference)

    if (imgwidth / imgheight >= threshold) {
        return "landscape"; // Landscape orientation
    } else if (imgheight / imgwidth >= threshold) {
        return "portrait"; // Portrait orientation
    } else {
        return "square"; // Square orientation
    }
}

async function getPptRecipientImages(itemID) {
    let attachments = await getAttachments(itemID);
    
    if (attachments.length === 0) {
        console.warn("⚠️ No attachments found, returning empty array.");
        return [];
    }

    let imagePromises = attachments.map(async (file) => {
        let imageUrl = file.ServerRelativeUrl;
        return convertImageToBase64(imageUrl); // Returns a Promise
    });

    return Promise.all(imagePromises); // Waits for all conversions
}

async function convertImageToBase64(imageUrl) {
    try {
        const token = await getSharePointToken();
        const response = await fetch(
            `https://dxcportal.sharepoint.com/sites/ITOEECoreTeam/_api/Web/GetFileByServerRelativeUrl('${imageUrl}')/$value`, 
            {
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Accept": "application/json;odata=verbose"
                }
            }
        );

        if (!response.ok) throw new Error(`Fetch failed with status ${response.status}`);

        const blob = await response.blob();
        const fileExtension = imageUrl.split('.').pop().toLowerCase(); // Extract file type from URL
        const mimeTypes = {
            png: "image/png",
            jpg: "image/jpeg",
            jpeg: "image/jpeg",
            gif: "image/gif",
            bmp: "image/bmp"
        };

        const mimeType = mimeTypes[fileExtension] || "image/png"; // Default to PNG if unknown
        const reader = new FileReader();

        return new Promise((resolve) => {
            reader.onloadend = () => resolve(`data:${mimeType};base64,${reader.result.split(",")[1]}`);
            reader.readAsDataURL(blob);
        });

    } catch (error) {
        console.error("Error converting image to Base64:", error);
    }
}

function checkWordCount(text,limit) {
    let count = text.trim().length;
    // let count = text.trim().split(/\s+/).length;
    // console.log(`Word count: ${count}, Limit: ${limit}`);
    return count > limit;
}

function getTemplateType(entry){
    let texts = {
        Challenge: entry.Challenge,
        Impact: entry.Impact,
        Uniqueness: entry.Uniqueness,
        Help: entry.Help
    };
    console.log({
        Challenge: texts.Challenge.trim().length,//texts.Challenge.trim().split(/\s+/).length,
        Impact: texts.Impact.trim().length,//texts.Impact.trim().split(/\s+/).length,
        Uniqueness: texts.Uniqueness? texts.Uniqueness.trim().length: null,//texts.Uniqueness? texts.Uniqueness.trim().split(/\s+/).length: null,
        Help: texts.Help.trim().length,//texts.Help.trim().split(/\s+/).length
    })
    //word limits
    //single page without uniqueness
    let sp_limits = {
        Challenge: 510,//80,
        Impact: 510,//80,
        Help: 510//80
    }
    //single page with uniqueness
    let spu_limits = {
        Challenge: 510,//80,
        Impact: 240,//35,
        Uniqueness: 240,//35,
        Help: 510//80
    }
    //double page without uniqueness
    let dp_limits = {
        Challenge: 930,//160,
        Impact: 860,//160,
        Help: 1900//310
    }
    //double page with uniqueness (short)
    let dpu_short_limits = {
        Challenge: 930,//160,
        Impact: 370,//60,
        Uniqueness: 370,//60,
        Help: 1900//310
    }
    //double page with uniqueness (long)
    let dpu_long_limits = {
        Challenge: 930,//160,
        Impact: 860,//160,
        Uniqueness: 930,//160
        Help: 860,//160,
    }
    //triple page
    let tp_limits = {
        Challenge: 1900,
        Impact: 1900,
        Help: 1900,
    }
    //full page
    let full_limits = {
        Challenge: 1900,
        Impact: 1900,
        Help: 1900,
        Uniqueness: 1900
    }
    let pages = 1;
    let template = "sp";
    if (texts.Uniqueness && texts.Uniqueness.trim().length > 0) {
        template = "spu"; // Start with single page with uniqueness
        // console.log("Challenge > spu_short:", checkWordCount(texts.Challenge, spu_limits.Challenge));
        // console.log("Impact > spu_short:", checkWordCount(texts.Impact, spu_limits.Impact));
        // console.log("Uniqueness > spu_short:", checkWordCount(texts.Uniqueness, spu_limits.Uniqueness));
        // console.log("Help > spu_short:", checkWordCount(texts.Help, spu_limits.Help));
        if (checkWordCount(texts.Challenge, spu_limits.Challenge) ||
            checkWordCount(texts.Help, spu_limits.Help) ||
            checkWordCount(texts.Impact, spu_limits.Impact) ||
            checkWordCount(texts.Uniqueness, spu_limits.Uniqueness)) {
            pages = 2;
            template = "dpu_short"; // Upgrade to double page (short)
        }
        // console.log("Challenge > dpu_short:", checkWordCount(texts.Challenge, dpu_short_limits.Challenge));
        // console.log("Impact > dpu_short:", checkWordCount(texts.Impact, dpu_short_limits.Impact));
        // console.log("Uniqueness > dpu_short:", checkWordCount(texts.Uniqueness, dpu_short_limits.Uniqueness));
        // console.log("Help > dpu_short:", checkWordCount(texts.Help, dpu_short_limits.Help));

        if (checkWordCount(texts.Challenge, dpu_short_limits.Challenge) ||
            checkWordCount(texts.Impact, dpu_short_limits.Impact) ||
            checkWordCount(texts.Uniqueness, dpu_short_limits.Uniqueness) ||
            checkWordCount(texts.Help, dpu_short_limits.Help)) {
            template = "dpu_long"; // Upgrade to double page (long)
        }
        // console.log("Challenge > dpu_long:", checkWordCount(texts.Challenge, dpu_long_limits.Challenge));
        // console.log("Impact > dpu_long:", checkWordCount(texts.Impact, dpu_long_limits.Impact));
        // console.log("Uniqueness > dpu_long:", checkWordCount(texts.Uniqueness, full_limits.Uniqueness));
        // console.log("Help > dpu_long:", checkWordCount(texts.Help, dpu_long_limits.Help));

        if (template != "dpu_short" && (checkWordCount(texts.Challenge, dpu_long_limits.Challenge) ||
            checkWordCount(texts.Help, dpu_long_limits.Help) ||
            checkWordCount(texts.Impact, dpu_long_limits.Impact) ||
            checkWordCount(texts.Uniqueness, dpu_long_limits.Uniqueness))) {
            pages = 4;
            template = "full"; // Full page
        }
    } else {
        if (checkWordCount(texts.Challenge, sp_limits.Challenge) ||
            checkWordCount(texts.Help, sp_limits.Help) ||
            checkWordCount(texts.Impact, sp_limits.Impact)) {
            pages = 2;
            template = "dp"; // Double page without uniqueness
        }

        if (checkWordCount(texts.Challenge, dp_limits.Challenge) ||
            checkWordCount(texts.Help, dp_limits.Help) ||
            checkWordCount(texts.Impact, dp_limits.Impact)) {
            pages = 3;
            template = "tp"; // Triple page
        }
    }

    return { template:template, pages: pages };
    

}
function individualEntry(entry, template){
    // Create a new PowerPoint presentation
    const pptx = new PptxGenJS();
    pptx.setTitle("Amazing Stories.pptx");
    pptx.setAuthor("DXC ITOEE Core Team");
    pptx.setCompany("DXC Technology");

    switch(template){
        case "sp": //single page without uniqueness
            createSinglePageWithoutUniqueness(pptx, entry);
            break;
        case "spu": //single page with uniqueness
            createSinglePageWithUniqueness(pptx, entry);
            break;
        case "dp_short": //double page with uniqueness (short)
            createDoublePageWithUniquenessShort(pptx, entry);
            break;
        case "dp_long": //double page with uniqueness (long)
            createDoublePageWithUniquenessLong(pptx, entry);
            break;
        case "tp": //triple page
            createTriplePage(pptx, entry);
            break;
        case "full": //full page
            createFullPage(pptx, entry);
            break;
    }

}
