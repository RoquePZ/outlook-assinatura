
function create_user_info() {
    const mensagem = document.getElementById("message");

    try {
        const perfil = {
            name: document.getElementById("display_name").value.trim(),
            email: document.getElementById("email_id").value.trim(),
            job: document.getElementById("job_title").value.trim(),
            department: document.getElementById("department").value.trim(),
            phone: document.getElementById("phone_number").value.trim(),
            greeting: document.getElementById("greeting_text").value.trim(),
            pronoun: document.getElementById("preferred_pronoun").value.trim()
        };

        if (!perfil.name || !perfil.email) {
            mensagem.textContent = "Preencha o nome e o e-mail.";
            return;
        }

        if (!Office || !Office.context || !Office.context.roamingSettings) {
            mensagem.textContent = "O Outlook ainda não inicializou. Feche e abra o suplemento novamente.";
            return;
        }

        Office.context.roamingSettings.set("user_info", perfil);

        Office.context.roamingSettings.saveAsync(function (resultado) {
            if (resultado.status === Office.AsyncResultStatus.Succeeded) {
                mensagem.textContent = "Assinatura salva! Feche e abra um novo e-mail para testar.";
            } else {
                mensagem.textContent = "Erro ao salvar: " +
                    (resultado.error ? resultado.error.message : "erro não identificado");
                console.error("Erro ao salvar perfil:", resultado.error);
            }
        });
    } catch (erro) {
        mensagem.textContent = "Falha ao salvar: " + erro.message;
        console.error(erro);
    }
}

function reset_all_configuration() {
    const mensagem = document.getElementById("message");

    try {
        Office.context.roamingSettings.remove("user_info");

        Office.context.roamingSettings.saveAsync(function (resultado) {
            mensagem.textContent =
                resultado.status === Office.AsyncResultStatus.Succeeded
                    ? "Configuração removida."
                    : "Não foi possível remover a configuração.";
        });
    } catch (erro) {
        mensagem.textContent = "Falha ao limpar: " + erro.message;
        console.error(erro);
    }
}
