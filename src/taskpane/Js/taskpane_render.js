// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

let _display_name;
let _job_title;
let _phone_number;
let _email_id;
let _greeting_text;
let _preferred_pronoun;
let _message;
let _booking_link;
let _working_hours;

Office.initialize = function(reason)
{
  on_initialization_complete();
}

let pca; // We gebruiken één globale variabele voor de MSAL instantie

Office.onReady(async (info) => {
    if (typeof msal === 'undefined') {
        console.error("MSAL bibliotheek is nog niet geladen!");
        return;
    }
    // Nu pas kun je initializeNAA() aanroepen
    await initializeNAA();
});

async function initializeNAA() {
    const msalConfig = {
        auth: {
            clientId: "e918ad24-1435-4770-b576-3a17f2a8b25a", // Vul hier je echte Client ID in
            authority: "https://microsoftonline.com", // Moet de volledige URL zijn
            supportsNestedAppAuth: true
        }
    };

    // Initialiseer de globale pca variabele
    pca = await msal.createNestablePublicClientApplication(msalConfig);
    console.log("NAA Initialized");
}

async function getJobTitleWithNAA() {
    // Zorg dat NAA eerst geïnitialiseerd is
    if (!pca) await initializeNAA();

    const authRequest = {
        scopes: ["https://microsoft.com/User.Read"] // Gebruik de volledige Graph scope
    };

    try {
        // Bij NAA hoef je vaak geen account mee te geven aan acquireTokenSilent, 
        // de host (Outlook) weet al wie de gebruiker is.
        const response = await pca.acquireTokenSilent(authRequest);
        const accessToken = response.accessToken;

        // Roep de juiste Microsoft Graph endpoint aan
        const graphResponse = await fetch("https://microsoft.com", {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        
        const userData = await graphResponse.json();
        console.log("Functietitel:", userData.jobTitle);
        return userData.jobTitle;

    } catch (error) {
        console.error("Token aanvraag mislukt:", error);
        // Let op: acquireTokenPopup werkt niet in alle Outlook-omgevingen zonder Office Dialog API
    }
}

function on_initialization_complete()
{
	$(document).ready
	(
		function()
		{
      _output = $("textarea#output");
      _display_name = $("input#display_name");
      _email_id = $("input#email_id");
      _job_title = $("input#job_title");
      _phone_number = $("input#phone_number");
      _greeting_text = $("input#greeting_text");
      _preferred_pronoun = $("input#preferred_pronoun");
      _work_address = $("input#work_address");
      _booking_link = $("input#booking_link");
      _working_hours = $("input#working_hours")
      _message = $("p#message");

      prepopulate_from_userprofile();
      load_saved_user_info();
		}
	);
}

function prepopulate_from_userprofile()
{
  _display_name.val(Office.context.mailbox.userProfile.displayName);
  _email_id.val(Office.context.mailbox.userProfile.emailAddress);
  (async () => {
    _job_title.val(await getJobTitleWithNAA());
  })()
  
}

function load_saved_user_info()
{
  let user_info_str = localStorage.getItem('user_info');
  if (!user_info_str)
  {
    user_info_str = Office.context.roamingSettings.get('user_info');
  }

  if (user_info_str)
  {
    const user_info = JSON.parse(user_info_str);

    _display_name.val(user_info.name);
    _email_id.val(user_info.email);
    _job_title.val(user_info.job);
    _phone_number.val(user_info.phone);
    _greeting_text.val(user_info.greeting);
    _preferred_pronoun.val(user_info.pronoun);
    _work_address.val(user_info.work_address);
    _booking_link.val(user_info.booking_link);
    _working_hours.val(user_info.working_hours);
  }
}

function display_message(msg)
{
  _message.text(msg);
}

function clear_message()
{
  _message.text("");
}

function is_not_valid_text(text)
{
  return text.length <= 0;
}

function is_not_valid_email_address(email_address)
{
  let email_address_regex = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/;
  return is_not_valid_text(email_address) || !(email_address_regex.test(email_address));
}

function form_has_valid_data(name, email)
{
  if (is_not_valid_text(name))
  {
    display_message("Please enter a valid name.");
    return false;
  }

  if (is_not_valid_email_address(email))
  {
    display_message("Please enter a valid email address.");
    return false;
  }

  return true;
}

function navigate_to_taskpane_assignsignature()
{
  window.location.href = 'assignsignature.html';
}

function create_user_info()
{
  let name = _display_name.val().trim();
  let email = _email_id.val().trim();

  clear_message();

  if (form_has_valid_data(name, email))
  {
    clear_message();

    let user_info = {};

    user_info.name = name;
    user_info.email = email;
    user_info.job = _job_title.val().trim();
    user_info.phone = _phone_number.val().trim();
    user_info.greeting = _greeting_text.val().trim();
    user_info.pronoun = _preferred_pronoun.val().trim();
    user_info.work_address = _work_address.val().trim();
    user_info.booking_link = _booking_link.val().trim();
    user_info.working_hours = _working_hours.val().trim();

    console.log(user_info);
    localStorage.setItem('user_info', JSON.stringify(user_info));
    navigate_to_taskpane_assignsignature();
  }
}

function clear_all_fields()
{
  _display_name.val("");
  _email_id.val("");
  _job_title.val("");
  _phone_number.val("");
  _greeting_text.val("");
  _preferred_pronoun.val("");
  _work_address.val("");
  _working_hours.val("");
  _booking_link.val("");
}

function clear_all_localstorage_data()
{
  localStorage.removeItem('user_info');
  localStorage.removeItem('newMail');
  localStorage.removeItem('reply');
  localStorage.removeItem('forward');
  localStorage.removeItem('override_olk_signature');
}

function clear_roaming_settings()
{
  Office.context.roamingSettings.remove('user_info');
  Office.context.roamingSettings.remove('newMail');
  Office.context.roamingSettings.remove('reply');
  Office.context.roamingSettings.remove('forward');
  Office.context.roamingSettings.remove('override_olk_signature');

  Office.context.roamingSettings.saveAsync
  (
    function (asyncResult)
    {
      console.log("clear_roaming_settings - " + JSON.stringify(asyncResult));

      let message = "All settings reset successfully! This add-in won't insert any signatures. You can close this pane now.";
      if (asyncResult.status === Office.AsyncResultStatus.Failed)
      {
        message = "Failed to reset. Please try again.";
      }

      display_message(message);
    }
  );
}

function reset_all_configuration()
{
  clear_all_fields();
  clear_all_localstorage_data();
  clear_roaming_settings();
}
