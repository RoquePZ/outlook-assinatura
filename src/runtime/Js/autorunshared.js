/* Assinatura automática PZ com banner personalizado */

function checkSignature(event) {
    let concluido = false;

    function finalizar() {
        if (!concluido) {
            concluido = true;
            event.completed();
        }
    }

    try {
        const mailbox = Office.context.mailbox;
        const salvo =
            Office.context.roamingSettings.get("user_info") || {};
        const perfil = mailbox.userProfile || {};

        const userInfo = {
            name: salvo.name || perfil.displayName || "",
            email: salvo.email || perfil.emailAddress || "",
            job: salvo.job || "",
            department: salvo.department || ""
        };

        if (!userInfo.name && !userInfo.email) {
            console.warn("Perfil da assinatura não configurado.");
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
        const fundoUrl =
            "https://roquepz.github.io/outlook-assinatura/fundo-assinatura.png";

        const imagem = new Image();

        await new Promise((resolve, reject) => {
            imagem.onload = resolve;
            imagem.onerror = () =>
                reject(new Error("Não foi possível carregar o banner."));
            imagem.src = fundoUrl;
        });

        const canvas = document.createElement("canvas");
        canvas.width = imagem.naturalWidth;
        canvas.height = imagem.naturalHeight;

        const ctx = canvas.getContext("2d");
        ctx.drawImage(imagem, 0, 0);

        const escalaX = canvas.width / 2048;
        const escalaY = canvas.height / 683;

        const x = 995 * escalaX;
        const y = 240 * escalaY;
        const espaco = 48 * escalaY;

        ctx.fillStyle = "#ffffff";
        ctx.textBaseline = "top";
        ctx.font = `bold ${Math.round(34 * escalaY)}px Arial`;
        ctx.fillText(userInfo.name || "", x, y);

        ctx.font = `${Math.round(29 * escalaY)}px Arial`;
        ctx.fillText(userInfo.department || "", x, y + espaco);
        ctx.fillText(userInfo.job || "", x, y + espaco * 2);

        ctx.fillText(
            userInfo.email || "",
            x + 70 * escalaX,
            y + espaco * 3
        );

        // Converte o banner personalizado para Base64.
        const dataUrl = canvas.toDataURL("image/png");
        const base64 = dataUrl.substring(dataUrl.indexOf(",") + 1);

        // Incorpora a imagem ao e-mail.
        Office.context.mailbox.item.addFileAttachmentFromBase64Async(
            base64,
            "assinatura-pz.png",
            { isInline: true },
            function (resultado) {
                if (resultado.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error("Erro ao incorporar imagem:", resultado.error);
                    eventObj.completed();
                    return;
                }

                const assinatura = `
                    <p style="margin:0;padding:0;">
                      <img
                        src="cid:assinatura-pz.png"
                        alt="Assinatura PZ Advogados"
                        width="700"
                        style="display:block;width:700px;height:auto;border:0;">
                    </p>
                `;

                Office.context.mailbox.item.body.setSignatureAsync(
                    assinatura,
                    { coercionType: Office.CoercionType.Html },
                    function (resultadoAssinatura) {
                        if (resultadoAssinatura.status !==
                            Office.AsyncResultStatus.Succeeded) {
                            console.error(
                                "Erro ao inserir assinatura:",
                                resultadoAssinatura.error
                            );
                        } else {
                            console.log(
                                "Assinatura PZ com banner inserida."
                            );
                        }

                        eventObj.completed();
                    }
                );
            }
        );
    } catch (error) {
        console.error("Erro ao gerar assinatura:", error);
        eventObj.completed();
    }
}

Office.actions.associate("checkSignature", checkSignature);
