
/*
 * Perfil do usuário via Microsoft Graph.
 * Execute carregarPerfilGraph() a partir de um botão
 * clicado pelo usuário no painel do complemento.
 */

const msalConfig = {
    auth: {
        clientId: "d4744baf-fdd4-45cd-8f1a-29a5f83f7e16",
        authority: "https://login.microsoftonline.com/3a967584-1413-4050-803c-4cf28a46cffe",
        redirectUri: "https://roquepz.github.io/outlook-assinatura/"
    },
    cache: {
        cacheLocation: "localStorage"
    }
};

const graphScopes = ["User.Read"];

async function carregarPerfilGraph() {
    try {
        const msalApp = new msal.PublicClientApplication(msalConfig);
        await msalApp.initialize();

        let account = msalApp.getAllAccounts()[0];

        if (!account) {
            const login = await msalApp.loginPopup({
                scopes: graphScopes
            });
            account = login.account;
        }

        let token;

        try {
            const resultado = await msalApp.acquireTokenSilent({
                scopes: graphScopes,
                account: account
            });
            token = resultado.accessToken;
        } catch {
            const resultado = await msalApp.acquireTokenPopup({
                scopes: graphScopes,
                account: account
            });
            token = resultado.accessToken;
        }

        const resposta = await fetch(
            "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,department",
            {
                headers: {
                    Authorization: "Bearer " + token
                }
            }
        );

        if (!resposta.ok) {
            throw new Error("Microsoft Graph retornou HTTP " + resposta.status);
        }

        const usuario = await resposta.json();

        const perfil = {
            name: usuario.displayName || "",
            email: usuario.mail || usuario.userPrincipalName || "",
            job: usuario.jobTitle || "",
            department: usuario.department || ""
        };

        Office.context.roamingSettings.set("user_info", perfil);

        await new Promise((resolve, reject) => {
            Office.context.roamingSettings.saveAsync(resultado => {
                if (resultado.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                } else {
                    reject(new Error("Não foi possível salvar o perfil no Outlook."));
                }
            });
        });

        return perfil;
    } catch (erro) {
        console.error("Falha ao consultar o perfil:", erro);
        throw erro;
    }
}
