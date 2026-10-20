/* Assinatura automática PZ - versão de recuperação */

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

        insert_auto_signature(userInfo, finalizar);
    } catch (error) {
        console.error("Erro ao preparar assinatura:", error);
        finalizar();
    }
}

function insert_auto_signature(userInfo, finalizar) {
    const esc = value => String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");

    const nome = esc(userInfo.name);
    const email = esc(userInfo.email);
    const cargo = esc(userInfo.job);
    const departamento = esc(userInfo.department);

    const assinatura = `
      <table cellpadding="0" cellspacing="0" border="0"
        style="font-family:Arial,sans-serif;font-size:10pt;color:#333;">
        <tr>
          <td style="padding-right:14px;vertical-align:middle;">
            <img src="https://roquepz.github.io/outlook-assinatura/Logo%20PZ.png"
              alt="Logo PZ" width="100"
              style="display:block;width:100px;height:auto;border:0;">
          </td>
          <td style="border-left:2px solid #d7d7d7;padding-left:14px;vertical-align:middle;">
            <strong style="font-size:11pt;color:#222;">${nome}</strong>
            ${departamento ? `<br>${departamento}` : ""}
            ${cargo ? `<br>${cargo}` : ""}
            ${email ? `<br><a href="mailto:${email}" style="color:#0563c1;">${email}</a>` : ""}
          </td>
        </tr>
      </table>`;

    Office.context.mailbox.item.body.setSignatureAsync(
        assinatura,
        { coercionType: Office.CoercionType.Html },
        function (result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Erro ao inserir assinatura:", result.error);
            } else {
                console.log("Assinatura de recuperação inserida.");
            }
            finalizar();
        }
    );
}

Office.actions.associate("checkSignature", checkSignature);
