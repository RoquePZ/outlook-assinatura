
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
    let finalizado = false;

    function finalizar() {
        if (!finalizado) {
            finalizado = true;

            if (
                eventObj &&
                typeof eventObj.completed === "function"
            ) {
                eventObj.completed();
            }
        }
    }

    function inserirTextoAlternativo(motivo) {
        console.error(
            "Assinatura em imagem não concluída:",
            motivo
        );

        const nome = escapeHtml(userInfo.name || "");
        const cargo = escapeHtml(userInfo.job || "");
        const departamento = escapeHtml(
            userInfo.department || ""
        );
        const email = escapeHtml(userInfo.email || "");

        const html =
            "<div style='font-family:Arial,sans-serif;font-size:10pt'>" +
            "<strong>" + nome + "</strong><br>" +
            (cargo ? cargo + "<br>" : "") +
            (departamento ? departamento + "<br>" : "") +
            (
                email
                    ? "<a href='mailto:" + email + "'>" +
                      email + "</a>"
                    : ""
            ) +
            "</div>";

        Office.context.mailbox.item.body.setSignatureAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            function (result) {
                if (
                    result.status !==
                    Office.AsyncResultStatus.Succeeded
                ) {
                    console.error(
                        "Erro na assinatura alternativa:",
                        result.error
                    );
                }

                finalizar();
            }
        );
    }

    try {
        const imagem = new Image();

        imagem.crossOrigin = "anonymous";

        imagem.onload = function () {
            try {
                const canvas = document.createElement("canvas");

                const larguraOriginal = imagem.naturalWidth;
                const alturaOriginal = imagem.naturalHeight;

                // Analisa a imagem para remover as faixas quase brancas
                // superior e inferior sem alterar a largura do banner.
                const canvasAnalise = document.createElement("canvas");
                canvasAnalise.width = larguraOriginal;
                canvasAnalise.height = alturaOriginal;

                const ctxAnalise = canvasAnalise.getContext("2d", {
                    willReadFrequently: true
                });

                if (!ctxAnalise) {
                    inserirTextoAlternativo(
                        "Não foi possível analisar a imagem do banner."
                    );
                    return;
                }

                ctxAnalise.drawImage(imagem, 0, 0);

                const pixels = ctxAnalise.getImageData(
                    0, 0, larguraOriginal, alturaOriginal
                ).data;

                function linhaTemConteudo(yLinha) {
                    // Amostra a linha horizontalmente para localizar
                    // a faixa colorida, ignorando o fundo quase branco.
                    const passo = 8;

                    for (
                        let px = 0;
                        px < larguraOriginal;
                        px += passo
                    ) {
                        const indice =
                            (yLinha * larguraOriginal + px) * 4;

                        const r = pixels[indice];
                        const g = pixels[indice + 1];
                        const b = pixels[indice + 2];
                        const a = pixels[indice + 3];

                        if (
                            a > 10 &&
                            (r < 235 || g < 235 || b < 235)
                        ) {
                            return true;
                        }
                    }

                    return false;
                }

                let topoRecorte = 0;
                let baseRecorte = alturaOriginal - 1;

                while (
                    topoRecorte < alturaOriginal - 1 &&
                    !linhaTemConteudo(topoRecorte)
                ) {
                    topoRecorte++;
                }

                while (
                    baseRecorte > topoRecorte &&
                    !linhaTemConteudo(baseRecorte)
                ) {
                    baseRecorte--;
                }

                // Mantém uma pequena margem de segurança nas bordas.
                const margem = Math.round(alturaOriginal * 0.005);

                topoRecorte = Math.max(
                    0,
                    topoRecorte - margem
                );

                baseRecorte = Math.min(
                    alturaOriginal - 1,
                    baseRecorte + margem
                );

                const alturaRecortada =
                    baseRecorte - topoRecorte + 1;

                canvas.width = larguraOriginal;
                canvas.height = alturaRecortada;

                const ctx = canvas.getContext("2d");

                if (!ctx) {
                    inserirTextoAlternativo(
                        "Não foi possível criar o canvas."
                    );
                    return;
                }

                // Desenha somente a área do banner, sem as margens brancas.
                ctx.drawImage(
                    imagem,
                    0,
                    topoRecorte,
                    larguraOriginal,
                    alturaRecortada,
                    0,
                    0,
                    larguraOriginal,
                    alturaRecortada
                );

                // Dados do usuário.
                const nome = String(userInfo.name || "");
                const cargo = String(userInfo.job || "");

                const departamento = String(
                    userInfo.department || ""
                );

                const email = String(userInfo.email || "");

                // Mantém a posição horizontal atual do texto.
                const x = 1080;
                const largura = 400;

                // Compensa o recorte para preservar a posição vertical.
                let y = 285 - topoRecorte;

                // Nome.
                ctx.textAlign = "left";
                ctx.textBaseline = "middle";
                ctx.fillStyle = "#FFFFFF";
                ctx.font = "bold 34px Arial";

                ctx.fillText(nome, x, y, largura);
                y += 48;

                // Cargo.
                ctx.fillStyle = "#F7B875";
                ctx.font = "26px Arial";

                if (cargo) {
                    ctx.fillText(cargo, x, y, largura);
                    y += 38;
                }

                // Departamento.
                ctx.fillStyle = "#FFFFFF";
                ctx.font = "24px Arial";

                if (departamento) {
                    ctx.fillText(
                        departamento,
                        x,
                        y,
                        largura
                    );

                    y += 36;
                }

                // E-mail.
                if (email) {
                    ctx.fillText(email, x, y, largura);
                }

                // Converte o banner finalizado em Base64.
                const base64 = canvas
                    .toDataURL("image/png")
                    .split(",")[1];

                const item = Office.context.mailbox.item;

                // Adiciona a imagem como anexo embutido.
                item.addFileAttachmentFromBase64Async(
                    base64,
                    "assinatura-pz.png",
                    { isInline: true },
                    function (resultadoAnexo) {
                        if (
                            resultadoAnexo.status !==
                            Office.AsyncResultStatus.Succeeded
                        ) {
                            inserirTextoAlternativo(
                                resultadoAnexo.error
                            );
                            return;
                        }

                        // Insere a imagem proporcionalmente no Outlook.
                        const html =
                            '<img src="cid:assinatura-pz.png" ' +
                            'alt="Assinatura PZ" width="1024" ' +
                            'style="display:block;width:100%;' +
                            'max-width:1024px;height:auto;' +
                            'border:0;margin:0;padding:0;">';

                        item.body.setSignatureAsync(
                            html,
                            {
                                coercionType:
                                    Office.CoercionType.Html
                            },
                            function (resultadoCorpo) {
                                if (
                                    resultadoCorpo.status !==
                                    Office.AsyncResultStatus.Succeeded
                                ) {
                                    console.error(
                                        "Erro ao inserir a imagem:",
                                        resultadoCorpo.error
                                    );
                                } else {
                                    console.log(
                                        "Imagem da assinatura inserida."
                                    );
                                }

                                finalizar();
                            }
                        );
                    }
                );

            } catch (erro) {
                inserirTextoAlternativo(erro);
            }
        };

        imagem.onerror = function () {
            inserirTextoAlternativo(
                "O banner não pôde ser carregado."
            );
        };

        imagem.src =
            "https://roquepz.github.io/outlook-assinatura/fundo-assinatura.png";

    } catch (erro) {
        inserirTextoAlternativo(erro);
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

Office.actions.associate(
    "checkSignature",
    checkSignature
);
