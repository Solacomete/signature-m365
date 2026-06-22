Office.onReady();

function insertSignature(event) {
    // 1. Récupérer l'email de la personne qui tape le message
    const userEmail = Office.context.mailbox.userProfile.emailAddress;

    // 2. HTML de test codé en dur (en attendant l'API)
    const testSignature = `
        <br><br>
        <div style="font-family: Arial, sans-serif; font-size: 10pt; color: #333;">
            <p><strong>Testeur Interne</strong></p>
            <p>Direction des Systèmes d'Information</p>
            <p>${userEmail}</p>
            <img src="https://solacomete.github.io/signature-m365/icon-64.png" alt="Bannière Test">
        </div>
    `;

    // 3. Insérer la signature
    Office.context.mailbox.item.body.setSignatureAsync(
        testSignature,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Erreur : " + asyncResult.error.message);
            }
            // 4. Clôturer l'événement (Obligatoire pour que le mail se charge)
            event.completed();
        }
    );
}

// 5. Exposer la fonction à Outlook
Office.actions.associate("insertSignature", insertSignature);
