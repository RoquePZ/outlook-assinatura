
/* Assinatura automática PZ no Outlook Web */

function insert_auto_signature(composeType, userInfo, eventObj) {
    let finalizado = false;

    function finalizar() {
        if (!finalizado) {
            finalizado = true;

            if (eventObj && typeof eventObj.completed === "function") {
                eventObj.completed();
            }
        }
    }

    function inserirAssinatura(html) {
        Office.context.mailbox.item.body.setSignatureAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error("Erro ao inserir assinatura:", result.error);
                } else {
                    console.log("Assinatura PZ inserida.");
                }

                finalizar();
            }
        );
    }

    try {
        const nome = String(userInfo.name || "");
        const cargo = String(userInfo.job || "");
        const departamento = String(userInfo.department || "");
        const email = String(userInfo.email || "");

        const imagem = new Image();
        imagem.crossOrigin = "anonymous";
        imagem.src =
            "https://roquepz.github.io/outlook-assinatura/fundo-assinatura.png";

        imagem.onload = function () {
            try {
                const canvas = document.createElement("canvas");
                canvas.width = imagem.naturalWidth;
                canvas.height = imagem.naturalHeight;

                const ctx = canvas.getContext("2d");
                ctx.drawImage(imagem, 0, 0);

                // Área central livre do banner.
                const x = 990;
                const largura = 440;
                let y = 285;

                ctx.textAlign = "left";
                ctx.textBaseline = "middle";
                ctx.fillStyle = "#FFFFFF";
                ctx.font = "bold 34px Arial";
                ctx.fillText(nome, x, y, largura);
                y += 48;

                ctx.fillStyle = "#F7B875";
                ctx.font = "26px Arial";

                if (cargo) {
                    ctx.fillText(cargo, x, y, largura);
                    y += 38;
                }

                ctx.fillStyle = "#FFFFFF";
                ctx.font = "24px Arial";

                if (departamento) {
                    ctx.fillText(departamento, x, y, largura);
                    y += 36;
                }

                if (email) {
                    ctx.fillText(email, x, y, largura);
                }

                const base64 = canvas.toDataURL("image/png")
                    .split(",")[1];

                Office.context.mailbox.item.addFileAttachmentFromBase64Async(
                    base64,
                    "assinatura-pz.png",
                    { isInline: true },
                    function (result) {
                        if (
                            result.status !==
                            Office.AsyncResultStatus.Succeeded
                        ) {
                            console.error(
                                "Erro ao adicionar imagem:",
                                result.error
                            );

                            // Alternativa para não deixar a assinatura vazia.
                            inserirAssinatura(
                                "<div style='font-family:Arial,sans-serif'>" +
                                "<strong>" + escapeHtml(nome) + "</strong><br>" +
                                escapeHtml(cargo) + "<br>" +
                                escapeHtml(departamento) + "<br>" +
                                escapeHtml(email) +
                                "</div>"
                            );
                            return;
                        }

                        const assinatura =
                            '<img src="cid:assinatura-pz.png" ' +
                            'alt="Assinatura PZ" width="1024" ' +
                            'style="display:block;width:1024px;' +
                            'max-width:100%;height:auto;border:0;">';

                        inserirAssinatura(assinatura);
                    }
                );
            } catch (error) {
                console.error("Erro ao gerar imagem:", error);
                finalizar();
            }
        };

        imagem.onerror = function () {
            console.error("Não foi possível carregar o banner.");
            finalizar();
        };
    } catch (error) {
        console.error("Erro na assinatura automática:", error);
        finalizar();
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
