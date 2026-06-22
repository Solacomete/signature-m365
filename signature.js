Office.onReady();

function insertSignature(event) {
    const userEmail = Office.context.mailbox.userProfile.emailAddress;

    fetch(`https://URL_DE_TON_API_CLOUD/api/signature?email=${userEmail}`)
        .then(response => response.json())
        .then(data => {
            if (data.html) {
                Office.context.mailbox.item.body.setSignatureAsync(
                    data.html,
                    { coercionType: Office.CoercionType.Html },
                    function (asyncResult) {
                        event.completed();
                    }
                );
            } else {
                event.completed();
            }
        })
        .catch(error => {
            event.completed();
        });
}

Office.actions.associate("insertSignature", insertSignature);