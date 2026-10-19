
/* Evento de criação de nova mensagem do Outlook */

function checkSignature(event) {
    console.log("EVENTO checkSignature EXECUTADO");    
    let concluido = false;

    function finalizar() {
        if (!concluido) {
            concluido = true;
            event.completed();
        }
    }

    try {
        const mailbox = Office.context.mailbox;
        const perfilSalvo =
            Office.context.roamingSettings.get("user_info") || {};

        const perfilOutlook = mailbox.userProfile || {};

        // Preserva os dados cadastrados pelo usuário.
        // Usa o perfil do Outlook como alternativa.
        const userInfo = {
            name: perfilSalvo.name || perfilOutlook.displayName || "",
            email: perfilSalvo.email ||
                   perfilOutlook.emailAddress || "",
            job: perfilSalvo.job || "",
            department: perfilSalvo.department || "",
            phone: perfilSalvo.phone || ""
        };

        if (!userInfo.name && !userInfo.email) {
            console.warn(
                "Perfil da assinatura não configurado."
            );
            finalizar();
            return;
        }

        if (typeof insert_auto_signature !== "function") {
            console.error(
                "A função insert_auto_signature não foi carregada."
            );
            finalizar();
            return;
        }

        insert_auto_signature(
            "new_message",
            userInfo,
            { completed: finalizar }
        );

    } catch (error) {
        console.error(
            "Erro ao preparar a assinatura:",
            error
        );
        finalizar();
    }
}

Office.actions.associate("checkSignature", checkSignature);
