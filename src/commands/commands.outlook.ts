/* global Office */

Office.onReady(() => {
  // Eventuele initialisatie logica
});

// Koppel de functie aan de actie uit het manifest
Office.actions.associate("insertSignature", insertSignature);

function insertSignature(event: Office.AddinCommands.Event): void {
  const item = Office.context.mailbox.item;

  // Gebruikersgegevens ophalen (Later uit te breiden met Entra ID / Graph)
  const userDisplayName = Office.context.mailbox.userProfile.displayName;
  const userEmail = Office.context.mailbox.userProfile.emailAddress;
  const jobTitle = "Functieomschrijving Placeholder "; 
  const mobileNumber = "+32 470 00 00 00"; 

  const signatureHtml = `
    <br/><br/>
    <hr style="border:none; border-top:1px solid #ccc; height:1px; margin:10px 0;" />
    <div style="font-family: Arial, sans-serif; font-size: 10pt; color: #333;">
      <strong>${userDisplayName}</strong><br />
      <span style="color: #666;">${jobTitle}</span><br />
      <span>GSM: ${mobileNumber}</span> | <span>E-mail: ${userEmail}</span>
    </div>
  `;

  item.body.setSignatureAsync(
    signatureHtml,
    { coercionType: Office.CoercionType.Html },
    (result: Office.AsyncResult<void>) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(result.error.message);
      }
      // Sluit het event correct af
      event.completed();
    }
  );
}