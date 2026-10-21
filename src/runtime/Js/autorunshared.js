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

function insert_auto_signature(composeType, userInfo, eventObj) {
    try {
        const nome = escapeHtml(userInfo.name || "");
        const email = escapeHtml(userInfo.email || "");
        const cargo = escapeHtml(userInfo.job || "");
        const departamento = escapeHtml(userInfo.department || "");

        const assinatura = `
          <div style="font-family:Arial,sans-serif;font-size:10pt;color:#333;">
            <strong>${nome}</strong><br>
            ${cargo ? cargo + "<br>" : ""}
            ${departamento ? departamento + "<br>" : ""}
            ${email ? '<a href="mailto:' + email + '">' + email + '</a>' : ""}
          </div>
        `;

        Office.context.mailbox.item.body.setSignatureAsync(
            assinatura,
            { coercionType: Office.CoercionType.Html },
            function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error("Não foi possível inserir a assinatura:", result.error);
                }

                if (eventObj && typeof eventObj.completed === "function") {
                    eventObj.completed();
                }
            }
        );
    } catch (error) {
        console.error("Erro ao inserir assinatura:", error);

        if (eventObj && typeof eventObj.completed === "function") {
            eventObj.completed();
        }
    }
}

function escapeHtml(value) {
    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

Office.actions.associate("checkSignature", checkSignature);
