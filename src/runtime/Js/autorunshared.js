
/* Inicialização do complemento de assinatura do Outlook */

Office.onReady(function () {
    // O Outlook carregou o Office.js.
});

function checkSignature(eventObj) {
    try {
        const mailbox = Office.context.mailbox;
        const profile = mailbox.userProfile;

        // Dados básicos disponíveis no perfil do Outlook.
        // Cargo e telefone serão obtidos em uma etapa posterior.
        const userInfo = {
            name: profile.displayName || "",
            email: profile.emailAddress || "",
            job: "",
            phone: ""
        };

        // Disponibiliza os dados para os outros arquivos do complemento.
        Office.context.roamingSettings.set("user_info", userInfo);

        if (typeof insert_auto_signature === "function") {
            insert_auto_signature("new_message", userInfo, eventObj);
        } else {
            console.error("A função insert_auto_signature não foi carregada.");
            eventObj.completed();
        }
    } catch (error) {
        console.error("Erro ao preparar a assinatura:", error);
        eventObj.completed();
    }
}

Office.actions.associate("checkSignature", checkSignature);
