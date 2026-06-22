Office.onReady();

function insertSignature(event) {
    // 1. Récupère l'adresse email de l'utilisateur actif dans Outlook
    const userEmail = Office.context.mailbox.userProfile.emailAddress;

    // 2. Interroge ton API Cloudflare Workers en lui passant l'email
    fetch(`https://signature-m365.louis-b15.workers.dev/?email=${userEmail}`)
        .then(response => response.json())
        .then(data => {
            if (data.html) {
                // 3. Injecte le HTML reçu de l'API dans le mail
                Office.context.mailbox.item.body.setSignatureAsync(
                    data.html,
                    { coercionType: Office.CoercionType.Html },
                    function (asyncResult) {
                        event.completed(); // Libère le thread Outlook
                    }
                );
            } else {
                event.completed();
            }
        })
        .catch(error => {
            console.error("Erreur API Signature:", error);
            event.completed();
        });
}

Office.actions.associate("insertSignature", insertSignature);
