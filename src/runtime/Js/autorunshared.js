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

        const userInfo = {
            name: perfilSalvo.name || perfilOutlook.displayName || "",
            email: perfilSalvo.email || perfilOutlook.emailAddress || "",
            job: perfilSalvo.job || "",
            department: perfilSalvo.department || "",
            phone: perfilSalvo.phone || ""
        };

        if (!userInfo.name && !userInfo.email) {
            console.warn("Perfil da assinatura não configurado.");
            finalizar();
            return;
        }

        if (typeof insert_auto_signature !== "function") {
            console.error("A função insert_auto_signature não foi carregada.");
            finalizar();
            return;
        }

        insert_auto_signature("new_message", userInfo, {
            completed: finalizar
        });

    } catch (error) {
        console.error("Erro ao preparar a assinatura:", error);
        finalizar();
    }
}

function insert_auto_signature(composeType, userInfo, eventObj) {
    try {
        const nome = escapeHtml(userInfo.name || "");
        const email = escapeHtml(userInfo.email || "");
        const cargo = escapeHtml(userInfo.job || "");
        const departamento = escapeHtml(userInfo.department || "");

        const fundo =
            "https://roquepz.github.io/outlook-assinatura/fundo-assinatura.png";

        const assinatura = `
          <table role="presentation" cellpadding="0" cellspacing="0"
            border="0" width="700"
            style="width:700px;border-collapse:collapse;font-family:Arial,sans-serif;">
            <tr>
              <td background="${fundo}" bgcolor="#10243a"
                width="700" height="180" valign="middle"
                style="width:700px;height:180px;background-color:#10243a;
                background-image:url('${fundo}');background-repeat:no-repeat;
                background-position:center;background-size:cover;padding:18px 24px;">

                <table role="presentation" cellpadding="0" cellspacing="0"
                  border="0" width="100%" style="border-collapse:collapse;">
                  <tr>
                    <td valign="middle" width="155"
                      style="width:155px;padding-right:18px;">
                      <img src="https://roquepz.github.io/outlook-assinatura/Logo%20PZ.png"
                        alt="PZ Advogados" width="135"
                        style="display:block;width:135px;height:auto;border:0;">
                    </td>

                    <td valign="middle"
                      style="font-family:Arial,sans-serif;color:#ffffff;
                      font-size:12px;line-height:1.6;">
                      <strong style="font-size:16px;color:#ffffff;">
                        ${nome}
                      </strong><br>
                      ${departamento
                        ? `<span>${departamento}</span><br>` : ""}
                      ${cargo
                        ? `<span>${cargo}</span><br>` : ""}
                      ${email
                        ? `<a href="mailto:${email}"
                            style="color:#ffffff;text-decoration:none;">${email}</a>`
                        : ""}
                    </td>
                  </tr>
                </table>

              </td>
            </tr>
          </table>
        `;

        Office.context.mailbox.item.body.setSignatureAsync(
            assinatura,
            { coercionType: Office.CoercionType.Html },
            function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error(
                        "Não foi possível inserir a assinatura:",
                        result.error
                    );
                } else {
                    console.log("Assinatura inserida com sucesso.");
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
