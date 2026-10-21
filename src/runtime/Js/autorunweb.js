/*
 * Assinatura automática PZ no Outlook Web.
 */

function insert_auto_signature(composeType, userInfo, eventObj) {
    try {
        const nome = escapeHtml(userInfo.name || "");
        const email = escapeHtml(userInfo.email || "");
        const cargo = escapeHtml(userInfo.job || "");
        const departamento = escapeHtml(userInfo.department || "");
        const telefone = escapeHtml(userInfo.phone || "");

        const logo =
            "https://roquepz.github.io/outlook-assinatura/Logo%20PZ.png";

        
const assinatura = `
  <table cellpadding="0" cellspacing="0" border="0"
         style="font-family:Arial,sans-serif;font-size:10pt;color:#333;">
    <tr>
      <td style="padding-right:14px;vertical-align:middle;">
        <img src="${logo}"
             alt="Logo PZ"
             width="100"
             style="display:block;width:100px;height:auto;border:0;">
      </td>
      <td style="border-left:2px solid #d7d7d7;padding-left:14px;vertical-align:middle;">
        <strong style="font-size:11pt;color:#222;">${nome}</strong>
        ${departamento ? `<br>${departamento}` : ""}
        ${cargo ? `<br>${cargo}` : ""}
        ${email ? `<br><a href="mailto:${email}" style="color:#0563c1;">${email}</a>` : ""}
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
                        "Erro ao inserir assinatura:",
                        result.error
                    );
                } else {
                    console.log("Assinatura PZ inserida com sucesso.");
                }

                if (eventObj &&
                    typeof eventObj.completed === "function") {
                    eventObj.completed();
                }
            }
        );
    } catch (error) {
        console.error("Erro ao inserir assinatura:", error);

        if (eventObj &&
            typeof eventObj.completed === "function") {
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
