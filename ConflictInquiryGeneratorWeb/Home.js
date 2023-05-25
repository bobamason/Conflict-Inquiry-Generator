(function () {
    "use strict";

    let messageBanner;

    // field elements information
    // type and label values can be used to dynamically add the Html elements
    let fieldElementInfos;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(() => {
            // Initialize the notification mechanism and hide it
            const element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();


            // If not using Word 2016, warn user.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This add-in is not supprted on old versions of MS Word.");
                $('#generate-button').hide()
                return;
            }

            fieldElementInfos = [
                {
                    id: "input-atty-initials",
                    rows: 1,
                    type: "textarea",
                    label: "Atty Initials",
                    toReplace: "[Atty Initials]"
                },
                {
                    id: "input-client-name",
                    rows: 4,
                    type: "textarea",
                    label: "Client Name",
                    toReplace: "[Client Name]"
                },
                {
                    id: "input-co-defendant",
                    rows: 1,
                    type: "textarea",
                    label: "Co-defendant",
                    toReplace: "[Co-defendant]"
                },
                {
                    id: "input-adverse-party",
                    rows: 4,
                    type: "textarea",
                    label: "Adverse Party",
                    toReplace: "[Adverse Party]"
                },
                {
                    id: "input-related-party",
                    rows: 1,
                    type: "textarea",
                    label: "Related Party",
                    toReplace: "[Related Party]"
                },
                {
                    id: "select-matter-type",
                    rows: 1,
                    type: "select",
                    label: "Matter Type",
                    toReplace: "[Matter Type]",
                    valueList: ["",
                        "Bankruptcy-2",
                        "Business / Commercial Non-Litigation-4",
                        "Civil Rights - Non Employment - Defense-5",
                        "Business / Commercial / Contract Litigation-3",
                        "Criminal / Municipal Court / Juvenile Court-6",
                        "Estate Planning - Wills, Trust, Gurdianships-7",
                        "Estate Administration and Probate Litigation-8"
                    ]
                },
                {
                    id: "input-court",
                    rows: 1,
                    type: "textarea",
                    label: "Court",
                    toReplace: "[Court]"
                },
                {
                    id: "input-originating-atty",
                    rows: 1,
                    type: "textarea",
                    label: "Originating Atty",
                    toReplace: "[Originating Atty]"
                },
                {
                    id: "input-responsible-atty",
                    rows: 1,
                    type: "textarea",
                    label: "Responsible Atty",
                    toReplace: "[Responsible Atty]"
                },
                {
                    id: "input-billing-atty",
                    rows: 1,
                    type: "textarea",
                    label: "Billing Atty",
                    toReplace: "[Billing Atty]"
                },
                {
                    id: "input-dms-matter",
                    rows: 1,
                    type: "textarea",
                    label: "DMS Matter",
                    toReplace: "[DMS Matter]"
                },
                {
                    id: "input-other-atty-para",
                    rows: 1,
                    type: "textarea",
                    label: "Other Atty Paralegal",
                    toReplace: "[Other Atty Paralegal]"
                },
                {
                    id: "input-your-initials",
                    rows: 1,
                    type: "textarea",
                    label: "Your Initials",
                    toReplace: "[Your Initials]"
                }
            ];

            $("#template-description").text("This add-in generates a conflict inquiry email with the information filled out in the from below.");

            //$('#button-text').text("Create Email");
            $('#button-desc').text("Generate the Conflict Inquiry email");

            //$('#clear-button-text').text("Clear Form");
            $('#clear-button-desc').text("Clear the entire form");
            
            // Populate form
            findAllFields().then(availableFieldsList => {

                for (let i = 0; i < fieldElementInfos.length; i++) {
                    addFormInput(fieldElementInfos[i], availableFieldsList);
                }

                $('#generate-button').click(onGenerateClicked);

                $('#clear-button').click(onClearFormClicked);
            });
        });
    };

    // add input control for an field element info object and check if the field exists in the template
    function addFormInput(elementInfo, availableFieldsList) {
        const formDiv = $('#content-form');

        const div = $('<div></div>');
        formDiv.append(div);
        div.addClass('flex');

        const label = $('<label></label>');
        div.append(label);
        label.text(elementInfo.label);
        label.attr({ for: elementInfo.id });
        label.addClass('ms-font-m');
        label.addClass('ms-fontWeight-semilight');

        if (elementInfo.type === "textarea") {
            const textarea = $('<textarea/>');
            div.append(textarea);
            textarea.attr({
                id: elementInfo.id,
                rows: elementInfo.rows,
                name: elementInfo.id
            });
            textarea.addClass('ms-font-m');

        } else if (elementInfo.type === "select") {
            const selectBox = $('<select></select>');
            div.append(selectBox);
            selectBox.attr({
                id: elementInfo.id,
                type: 'text',
                name: elementInfo.id
            });
            selectBox.addClass('ms-font-m');

            populateSelectValues(elementInfo.id, elementInfo.valueList);
        }

        if (!availableFieldsList.includes(elementInfo.toReplace)) {
            const errorText = $('<p></p>');
            errorText.text(`Add ${elementInfo.toReplace} to the template to use the field above!`);
            errorText.addClass('ms-font-s');
            errorText.addClass('ms-fontColor-alert');
            div.append(errorText);
        }

        formDiv.append($('<br/>'));
    }

    function populateSelectValues(elementId, values) {
        if (values === undefined) return;
        values.forEach((v) => {
            $(`#${elementId}`).append($('<option></option>').text(v));
        });
    }

    async function onGenerateClicked() {

        let bodyText = await getBodyText();

        for (let i = 0; i < fieldElementInfos.length; i++) {

            tryCatch(async () => {
                const fieldInfo = fieldElementInfos[i];
                const toReplace = fieldInfo.toReplace;

                const fieldValue = $(`#${fieldInfo.id}`).prop('value');

                if (fieldValue) {

                    console.log(`#${fieldInfo.id}: Replacing field ${toReplace} with ${fieldValue}`);

                    bodyText = bodyText.replace(toReplace, fieldValue);

                    //bodyText = bodyText.replace("\[" + toReplace.substring(1, toReplace.length - 1) + "\]", fieldValue);
                } else {
                    console.log(`Empty input field ${toReplace}`);
                }

            });
        }

        openEmailWindow("", "Conflict of Interest", bodyText);
    }

    function onClearFormClicked() {
        for (let i = 0; i < fieldElementInfos.length; i++) {
            tryCatch(async () => {
                $(`#${fieldElementInfos[i].id}`).val("");
            });
        }
    }


    //// Don't use since before Office 365 this will overwrite the template
    //async function searchAndReplace(textToReplace, newText) {

    //    tryCatch(async () => {
    //        await Word.run(async (context) => {
    //            // Construct a wildcard expression and set matchWildcards to true in order to use wildcards.
    //            const results = context.document.body.search(textToReplace, { matchWildcards: false });
    //            results.load("text, font");

    //            await context.sync();

    //            // Let's traverse the search results and replace matches.
    //            for (let i = 0; i < results.items.length; i++) {
    //                //results.items[i].font.highlightColor = "red";
    //                results.items[i].insertText(newText, Word.InsertLocation.replace);
    //            }

    //            await context.sync();
    //        });
    //    });
    //}


    // Finds all fields in the document text marked with brackets example [Some Field]
    // returns an array of text field strings enclosed in brackets
    async function findAllFields() {
        const found = [];
        await Word.run(async (context) => {
                // Construct a wildcard expression and set matchWildcards to true in order to use wildcards.
                const results = context.document.body.search("[[]*[]]", { matchWildcards: true });
                results.load("text, font");

                await context.sync();

                // Let's traverse the search results and highlight matches.
                for (let i = 0; i < results.items.length; i++) {
                    //results.items[i].font.highlightColor = "red";
                    console.log("found tag: " + i + ", " + results.items[i].text);
                    const text = results.items[i].text;
                    found.push(text);
                }
        });

        return found;
    }

    async function getBodyText() {
        let text = "";

        await Word.run(async (context) => {
                const body = context.document.body;
                body.load("text");

                await context.sync();
            
                text = body.text;
        });

        return text;
    }

    // opens the system's default email app
    // recipents must be a comma separated string of addresses with no spaces!
    function openEmailWindow(recipients, subject, body) {
        const escapedSubject = encodeURI(subject).replace("+", "%20");

        const escapedBody = encodeURI(body).replace("+", "%20");

        const uri = `mailTo:${recipients}?subject=${escapedSubject}&body=${escapedBody}`;

        console.log("opening email uri: " + uri);

        window.open(uri);
    }


    async function tryCatch(callback) {
        try {
            await callback();
        } catch (error) {
            // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
            errorHandler(error);
        }
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
