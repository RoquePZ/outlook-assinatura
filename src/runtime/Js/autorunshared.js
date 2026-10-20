/* Assinatura automática personalizada PZ */

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

        const userInfo = {
            name: perfilSalvo.name || perfilOutlook.displayName || "",
            email: perfilSalvo.email || perfilOutlook.emailAddress || "",
            job: perfilSalvo.job || "",
            department: perfilSalvo.department || ""
        };

        if (!userInfo.name && !userInfo.email) {
            console.warn("Perfil da assinatura não configurado.");
            finalizar();
            return;
        }

        if (typeof gerarImagemAssinatura !== "function") {
            console.error("A função gerarImagemAssinatura não foi carregada.");
            finalizar();
            return;
        }

        insert_auto_signature("new_message", userInfo, {
            completed: finalizar
        });

    } catch (error) {
        console.error("Erro ao preparar assinatura:", error);
        finalizar();
    }
}

async function insert_auto_signature(composeType, userInfo, eventObj) {
    try {
        const dataUrl = await gerarImagemAssinatura(userInfo);
        console.log("Imagem gerada:", dataUrl.substring(0, 50));

        const assinatura = `
          <p style="margin:0;padding:0;">
            <img src="${dataUrl}"
                 alt="Assinatura PZ Advogados"
                 width="2048"
                 style="display:block;width:100%;max-width:700px;height:auto;border:0;">
          </p>
        `;

        Office.context.mailbox.item.body.setSignatureAsync(
            assinatura,
            { coercionType: Office.CoercionType.Html },
            function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error("Erro ao inserir assinatura:", result.error);
                } else {
                    console.log("Assinatura personalizada inserida.");
                }

                if (eventObj && eventObj.completed) {
                    eventObj.completed();
                }
            }
        );

    } catch (error) {
        console.error("Erro ao gerar a imagem da assinatura:", error);

        if (eventObj && eventObj.completed) {
            eventObj.completed();
        }
    }
}

Office.actions.associate("checkSignature", checkSignature);
