
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

        console.log("PZ: evento executado", userInfo);

        insert_auto_signature(userInfo, finalizar);
    } catch (e) {
        console.error("PZ: erro no evento", e);
        finalizar();
    }
}

function insert_auto_signature(userInfo, finalizar) {
    const fundoUrl =
        "https://roquepz.github.io/outlook-assinatura/fundo-assinatura.png";

    const imagem = new Image();

    imagem.onload = function () {
        try {
            const canvas = document.createElement("canvas");
            canvas.width = imagem.naturalWidth;
            canvas.height = imagem.naturalHeight;

            const ctx = canvas.getContext("2d");
            ctx.drawImage(imagem, 0, 0);

            const sx = canvas.width / 2048;
            const sy = canvas.height / 683;
            const x = 995 * sx;
            const y = 240 * sy;
            const espaco = 48 * sy;

            ctx.fillStyle = "#ffffff";
            ctx.textBaseline = "top";
            ctx.font = `bold ${Math.round(34 * sy)}px Arial`;
            ctx.fillText(userInfo.name || "", x, y);

            ctx.font = `${Math.round(29 * sy)}px Arial`;
            ctx.fillText(userInfo.department || "", x, y + espaco);
            ctx.fillText(userInfo.job || "", x, y + espaco * 2);
            ctx.fillText(userInfo.email || "", x + 70 * sx, y + espaco * 3);

            const base64 = canvas.toDataURL("image/png").split(",")[1];

            console.log("PZ: imagem gerada", canvas.width, canvas.height);

            Office.context.mailbox.item.addFileAttachmentFromBase64Async(
                base64,
                "assinatura-pz.png",
                { isInline: true },
                function (anexo) {
                    if (anexo.status !== Office.AsyncResultStatus.Succeeded) {
                        console.error("PZ: erro no anexo", anexo.error);
                        finalizar();
                        return;
                    }

                    const html =
                        '<img src="cid:assinatura-pz.png" ' +
                        'alt="Assinatura PZ Advogados" ' +
                        'width="700" style="display:block;width:700px;height:auto;border:0;">';

                    Office.context.mailbox.item.body.setSignatureAsync(
                        html,
                        { coercionType: Office.CoercionType.Html },
                        function (resultado) {
                            if (resultado.status !== Office.AsyncResultStatus.Succeeded) {
                                console.error("PZ: erro na assinatura", resultado.error);
                            } else {
                                console.log("PZ: assinatura inserida");
                            }
                            finalizar();
                        }
                    );
                }
            );
        } catch (e) {
            console.error("PZ: erro ao gerar banner", e);
            finalizar();
        }
    };

    imagem.onerror = function () {
        console.error("PZ: não foi possível carregar o banner", fundoUrl);
        finalizar();
    };

    imagem.src = fundoUrl;
}

Office.actions.associate("checkSignature", checkSignature);
