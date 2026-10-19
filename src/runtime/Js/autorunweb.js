```javascript
/*
 * Inserção automática da assinatura no Outlook.
 * Os dados são lidos das configurações do complemento.
 */

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
            {
                coercionType: Office.CoercionType.Html,
                asyncContext: eventObj
            },
            function (result) {
                console.log("Resultado da inserção:", result.status);

                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error(
                        "Erro ao inserir assinatura:",
                        result.error
                    );
                }

                const evento = result.asyncContext;
                if (evento && typeof evento.completed === "function") {
                    evento.completed();
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
```
