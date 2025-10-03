//global variables
const mistral_greetings = [
    "Fancy seeing you here!",
    "Well hello there!",
    "Look who it is!",
    "Hey, glad you dropped by!",
    "Ah, just the person I was hoping for!",
    "Hey hey! Ready to dive in?",
    "Welcome back, superstar!",
    "There you are — I was just thinking about you!",
    "Good to see you again!",
    "Hey, let\’s make something awesome!",
    "Oh hey, didn't see you there!",
    "Back for more? I like your style.",
    "You again? This must be fate.",
    "Just in time — I was getting bored.",
    "Ah, the legend returns!",
    "You always know when to show up.",
    "Let\’s make some magic, shall we?",
    "I was hoping you'd click that.",
    "You\’ve got great timing, as always.",
    "Let\’s do something brilliant together!"
];
const mistral_inputTooShort = [
  "Hmm, this seems a bit short.",
  "Want to add a little more detail?",
  "Looks like this could use some expansion.",
  "That’s a good start — care to elaborate?",
  "I think there’s more you could say here.",
  "This feels a bit light — want to flesh it out?",
  "You might want to add a few more thoughts.",
  "Let’s build this out a bit more.",
  "Could you give me a little more context?",
  "This might be clearer with a bit more info.",
  "I’m picking up the vibe, but it’s a bit thin.",
  "Let’s add some substance to this.",
  "Think you could expand on that?",
  "This could benefit from a few more lines.",
  "It’s a bit brief — want to go deeper?",
  "I’m sure there’s more to say here!",
  "Let’s give this some extra weight.",
  "Feels like it’s missing some key details.",
  "Want to round this out a bit more?",
  "You’re on the right track — now let’s build on it!"
];
const mistral_inputTooLong = [
  "Hmm, this seems a bit long-winded.",
  "Think we could trim this down a bit?",
  "That’s quite the novel — want to tighten it up?",
  "This might be clearer with fewer words.",
  "Let’s try to make this more concise.",
  "You’ve got a lot to say — maybe too much?",
  "This could use a little editing for brevity.",
  "Whew, that’s a lot! Want to simplify it?",
  "Let’s aim for clarity over length.",
  "You might lose your reader halfway through this.",
  "This is starting to feel like a monologue.",
  "Could we say the same thing with fewer words?",
  "Let’s trim the fat and keep the flavor.",
  "That’s a bit of a wall of text — want help condensing it?",
  "This might benefit from a more focused version.",
  "You’ve got the ideas — now let’s sharpen the delivery.",
  "Let’s make this punchier.",
  "A little editing could go a long way here.",
  "Want me to help you tighten this up?",
  "Let’s cut to the chase — I can help!"
];
const mistral_improvements = [
  "That's a solid draft. Want to improve it?",
  "Nice start! Want to make it even sharper?",
  "Looking good — ready to polish it up?",
  "This has potential. Want to refine it together?",
  "You're on the right track. Want to elevate it?",
  "Great bones here — shall we tighten it up?",
  "This works! Want to make it pop a bit more?",
  "Strong draft. Want to boost clarity or tone?",
  "You’ve got the idea — want help refining the delivery?",
  "This is promising. Want to enhance it?",
  "Solid effort! Want to make it even more impactful?",
  "You're close — want to fine-tune it?",
  "This could shine with a few tweaks. Want help?",
  "Nice flow! Want to sharpen the message?",
  "This is working — want to level it up?",
  "Great draft. Want to make it more concise?",
  "You’ve got something strong here. Want to polish it?",
  "This is nearly there. Want to refine it a bit?",
  "Good structure! Want to enhance the tone or clarity?",
  "Looks good! Want to explore a stronger version?"
];
const mistral_holdOn = [
  "Okay, hold on.",
  "Got it — one sec.",
  "Hang tight.",
  "Just a moment.",
  "Hold that thought.",
  "Give me a second.",
  "Alright, give me a moment.",
  "One moment, please.",
  "Let me check that.",
  "Working on it...",
  "Hold up a sec.",
  "Okay, let me think.",
  "Stand by...",
  "Just a sec!",
  "Alright, hang on.",
  "Let me pull that up.",
  "Give me a tick.",
  "Hold on, almost there.",
  "Okay, processing...",
  "Let me take a look."
];
const mistral_allDone = [
  "All set!",
  "Done and delivered!",
  "Here it is!",
  "Ready for you!",
  "Boom — there you have it!",
  "Your request, served fresh.",
  "As promised!",
  "Here’s what you asked for.",
  "Delivered, just like that.",
  "This one’s for you.",
  "Voila!",
  "Take a look at this!",
  "Hot off the press!",
  "Here’s the result!",
  "Mission accomplished.",
  "And… done!",
  "Right on cue.",
  "Here’s the magic.",
  "Freshly generated for you.",
  "Wrapped and ready!"
];


async function callMyAI(prompt) {
    console.log("Calling AI with prompt:", prompt);
    const response = await fetch("/api/hug", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ prompt })
    });

    if (!response.ok) {
        const errorText = await response.text();
        console.error("Server error:", errorText);
        return {error:errorText,message:"Sorry, something's wrong with my AI."};
    }

    const data = await response.json();
    console.log("AI response:", data.message);
    return data;
}

function mistralCheckDraft(element){
    const type = element.attr('contentType');
    let textarea = element.attr('textarea');
    let id = element.attr('id');
    let draft = $('#'+textarea).val().trim();
    let selector = '#'+id;
    let message = "Hey hey.";
    if(draft.length == 0){
        message = `Need help getting started with <b>${type}</b>? Just type the key points into the text area and we can work something out!`;
    }else if(draft.length < 200){
        message = mistral_inputTooShort[getRandomIndex(mistral_inputTooShort)] + ` Let's make it at least 200 characters long.
            <div class="d-flex align-items-center justify-content-center">
                <button type="button" class="btn btn-primary m-1 w-100 mistral-improve-draft" trigger="${selector}" textarea="${textarea}" contenttype="${type}">Lengthen</button>
            </div>`;
    }else if(draft.length > 1900){
        message = mistral_inputTooLong[getRandomIndex(mistral_inputTooLong)] + ` Let's make it less than 1900 characters long.
            <div class="d-flex align-items-center justify-content-center">
                <button type="button" class="btn btn-primary m-1 w-100 mistral-improve-draft" trigger="${selector}" textarea="${textarea}" contenttype="${type}">Shorten</button>
            </div>`;
    }else{
        message = `${mistral_improvements[getRandomIndex(mistral_improvements)]}
            <div class="d-flex align-items-center justify-content-center">
                <div class="d-flex flex-column align-items-center justify-content-center mx-1">
                    <button type="button" class="btn btn-primary m-1 w-100 mistral-improve-draft" trigger="${selector}" textarea="${textarea}" contenttype="${type}">Shorten</button>
                    <button type="button" class="btn btn-primary m-1 w-100 mistral-improve-draft" trigger="${selector}" textarea="${textarea}" contenttype="${type}">Lengthen</button>
                </div>
                <div class="d-flex flex-column align-items-center justify-content-center mx-1">
                    <button type="button" class="btn btn-primary m-1 w-100 mistral-improve-draft" trigger="${selector}" textarea="${textarea}" contenttype="${type}">Clean up</button>
                    <button type="button" class="btn btn-primary m-1 w-100 mistral-improve-draft" trigger="${selector}" textarea="${textarea}" contenttype="${type}">Rephrase</button>
                </div>
            </div>`;
    }
    callTippy(selector, message,"right")
}
async function mistralImproveDraft(textarea,intent,draft,type,trigger){
    let instructions=[];
    let lengthInstruction = "";
    const holdOn = mistral_holdOn[getRandomIndex(mistral_holdOn)];
    callTippy(trigger,holdOn,"right");
    switch (intent) {
        case "Shorten":
            lengthInstruction = "The new draft must be AT LEAST 200 characters long";
            break;
        case "Lengthen":
            lengthInstruction = "The new draft must be NO MORE THAN 1900 characters long";
            break;
        default:
            lengthInstruction = "The new draft must be BETWEEN 200 and 1900 characters long";
            break;
    }
    instructions.push(`You are helping improve a professional entry for: ${type}.`);
    instructions.push(`DO NOT include a character count in your response or a puppy dies.`);
    instructions.push(`Here is the current draft:"""${draft}"""`);
    instructions.push(`Please ${intent.toLowerCase()} the draft.`);
    instructions.push(`${lengthInstruction}, including all symbols and spaces.`);
    instructions.push(`Do not exceed this limit.`);
    instructions.push(`DO NOT include a character count in your response or a kitten dies.`);
    instructions.push(`Only respond with the revised draft.`);
    instructions.push(`Strictly use one paragraph only. Do not use bullet points.`);
    instructions.push(`Do not add a title.`)
    const prompt = instructions.join(' ');
    console.log(prompt)
    let response = await callMyAI(prompt);
    if (!response.error) {
        let cleanOutput = cleanMistralOutput(response.message.trim());

        let tries = 0;
        while ((cleanOutput.length > 1900 || cleanOutput.length < 200) && tries < 3) {
            response = await callMyAI(prompt);
            if (response.error) {
                callTippy(trigger, response.message, "right");
                return;
            }
            cleanOutput = cleanMistralOutput(response.message.trim());
            tries++;
        }

        textarea.val(cleanOutput);
        const doneMessage = mistral_allDone[getRandomIndex(mistral_allDone)];
        callTippy(trigger, doneMessage, "right");
        textarea.next('.text-count').text(`${cleanOutput.length} characters`);
    } else {
        callTippy(trigger, response.message, "right");
    }

}




function cleanMistralOutput(text) {
  return text
    // Matches: (1900 characters), (Exactly 1900 characters), (200-1900 characters), etc.
    .replace(/\(\s*(exactly|approximately|approx\.?|about|around)?\s*\d{1,3}(?:,\d{3})*(\s*[-–]\s*\d{1,3}(?:,\d{3})*)?\s*characters?\s*\)/gi, '')
    // Character count: 1234
    .replace(/character count\s*:\s*\d{1,3}(?:,\d{3})*/gi, '')
    // Total: 1,234 characters
    .replace(/total\s*:\s*\d{1,3}(?:,\d{3})*\s*characters?/gi, '')
    // Approx. 1234 chars, About 1,000 chars, etc.
    .replace(/(exactly|approximately|approx\.?|about|around)?\s*\d{1,3}(?:,\d{3})*(\s*[-–]\s*\d{1,3}(?:,\d{3})*)?\s*chars?/gi, '')
    // Collapse extra spaces
    .replace(/\s{2,}/g, ' ')
    .trim();
}





$(document).ready(function(){
    $(document)
    .on('mouseenter', '.mistral-button', function () {
        const greeting = mistral_greetings[getRandomIndex(mistral_greetings)];
        const el = this;

        if (el._tippy) {
            el._tippy.setContent(greeting);
            el._tippy.show();
        } else {
            tippy(el, {
            content: greeting,
            placement: 'right',
            arrow: true,
            trigger: 'manual'
            }).show();
        }
    })
    /* .on('mouseleave', '.mistral-button', function () {
        const el = this;
        if (el._tippy) {
            el._tippy.hide();
        }
    }) */
    .on('click','.mistral-button',function(){
        mistralCheckDraft($(this));
    })
    .on('click','.mistral-improve-draft', async function(){
        let textarea = $('#'+$(this).attr('textarea'));
        let intent = $(this).text();
        let draft = textarea.val().trim();
        let type = $(this).attr('contenttype');
        let trigger = $(this).attr('trigger');
        mistralImproveDraft(textarea,intent,draft,type,trigger)
        
    })

});