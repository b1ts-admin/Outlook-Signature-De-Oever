// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on Outlook on web, on Windows, and on Mac (new UI preview).

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */

function checkSignature(eventObj) {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  if (!user_info_str) {
    display_insight_infobar();
  } else {
    let user_info = JSON.parse(user_info_str);

    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync(
        {
          asyncContext: {
            user_info: user_info,
            eventObj: eventObj,
          },
        },
        function (asyncResult) {
          if (asyncResult.status === "succeeded") {
            insert_auto_signature(
              asyncResult.value.composeType,
              asyncResult.asyncContext.user_info,
              asyncResult.asyncContext.eventObj
            );
          }
        }
      );
    } else {
      // Appointment item. Just use newMail pattern
      let user_info = JSON.parse(user_info_str);
      insert_auto_signature("newMail", user_info, eventObj);
    }
  }
}

/**
 * For Outlook on Windows and on Mac only. Insert signature into appointment or message.
 * Outlook on Windows and on Mac can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
  addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Please set your signature with the Office Add-ins sample.",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Set signatures",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  if (template_name === "templateC") return get_template_C_info(user_info);
  return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */
function get_template_A_info(user_info) {
  const logoFileName = "sample-logo.png";
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += is_valid_data(user_info.booking_link) ? "<p>" + user_info.working_hours : "";
    str += is_valid_data(user_info.booking_link) ? "<br><a href=\"" + user_info.booking_link + "\"><span style=\"color:rgb(0,0,0);font-size:12pt;\"><picture><source srcset=\"https://ckbox.cloud/caab7c40545f8a590534/assets/kOtRSNxoIDv9/images/20.webp 20w\" sizes=\"(max-width: 20px) 100vw, 20px\" type=\"image/webp\"><img src=\"https://ckbox.cloud/caab7c40545f8a590534/assets/kOtRSNxoIDv9/images/20.png\" width=\"20\" height=\"20\"></picture></span><span style=\"color:rgb(0,120,212);font-size:12pt;\">Maak een afspraak voor een gesprek met mij</span></a>" : "";
    str += is_valid_data(user_info.booking_link) ? "</p>" : "";
    str += "<p>" + user_info.greeting + "</p>";
  }

  str += "<p style=\"color:#004259;font-size:medium;\">";
  str += "<strong>" + user_info.name + "</strong>";
  str += "<br/>";
  str += "<span style=\"color:#0073a5;font-size:small;\">" + user_info.job + "</span>";
  str += "</p>";
  str += "<span style=\"color:rgb(0,66,89);font-size:12.22px;\">" + user_info.work_address + "</span><br>";
  str += "<a style=\"color:#004259 !important';font-size:small;text-decoration:none;\" href=\"https://www.deoever.be\"><strong>www.deoever.be</strong></a>";
  str += "<table style=\"height:20pt;width:500pt;\">";
  str +=   "<colgroup><col style=\"width:10%;\"><col style=\"width:90%;\"></colgroup>";
  str +=   "<tbody>";
  str +=     "<tr>";
  str +=       "<td style=\"height:20pt;width:50pt;\">"
  str +=         "<a href=\"https://be.linkedin.com/company/vzw-de-oever\"><img src=\"https://imgmsgen.com/img/bookmark/ln.png\" width=\"20\" height=\"20\"> </a><a href=\"https://www.instagram.com/deoevervzw/\"><img src=\"https://imgmsgen.com/img/bookmark/it.png\" width=\"20\" height=\"20\"></a>";
  str +=       "</td>"
  str +=       "<td style=\"height:20pt;width:450pt;\">";
  str +=         "<span style=\"color:rgb(0,0,0);\"><strong>Partner Agentschap Opgroeien - Jeugdhulp</strong></span><br>"
  str +=         "<span style=\"color:rgb(0,115,165);\">0413.895.634 | RPR Ondernemingsrechtbank Antwerpen afdeling Hasselt</span>"
  str +=       "</td>";
  str +=     "</tr>";
  str +=   "</tbody>";
  str += "</table>";
  str += "<img style=\"width:500pt;\" src=\"https://www.deoever.be/wp-content/uploads/2024/09/banner-algemeen.png\"><br>";
  str += "<p style=\"text-align:justify;color: gray;font-size: xx-small\">";
  str +=   "Deze e-mail en zijn bijlagen zijn uitsluitend bestemd voor de geadresseerde(n) en strikt vertrouwelijk. Hun inhoud kan bij wet beschermd zijn. Indien de mail niet voor u bestemd is, is elke publicatie, reproductie, kopie, distributie of andere verspreiding of gebruik ervan ten strengste verboden. Als u deze boodschap per vergissing toegestuurd kreeg, gelieve de afzender onmiddellijk te verwittigen en de e-mail te vernietigen. Vzw De Oever besteedt de uiterste zorg aan de betrouwbaarheid en actualiteit van de gegevens die het verspreidt. Desalniettemin blijven fouten mogelijk, ook bij de transmissie van de gegevens. De overgebrachte informatie kan onderschept, gewijzigd of vernietigd zijn. Ze kan ook verloren gaan, te laat of onvolledig aankomen of een virus bevatten. Vzw De Oever aanvaardt bijgevolg geen enkele verantwoordelijkheid voor schade als gevolg van onjuistheden of van problemen veroorzaakt door of inherent aan het verspreiden van informatie via e-mail, evenals voor technische storingen en virussen.";
  str += "</p>";


  // return object with signature HTML, logo image base64 string, and filename to reference it with.
  return {
    signature: str,
    logoBase64:
      "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC",
    logoFileName: logoFileName,
  };
}

/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
  let str = get_template_B_str(user_info);

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
  let str = get_template_C_str(user_info);
  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
