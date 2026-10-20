/* Gera a imagem personalizada da assinatura PZ */

async function gerarImagemAssinatura(userInfo) {
    const fundoUrl =
        "https://roquepz.github.io/outlook-assinatura/fundo-assinatura.png";

    const imagemFundo = new Image();
    imagemFundo.crossOrigin = "anonymous";

    await new Promise((resolve, reject) => {
        imagemFundo.onload = resolve;
        imagemFundo.onerror = () =>
            reject(new Error("Não foi possível carregar o fundo."));
        imagemFundo.src = fundoUrl;
    });

    const canvas = document.createElement("canvas");
    canvas.width = imagemFundo.naturalWidth;
    canvas.height = imagemFundo.naturalHeight;

    const ctx = canvas.getContext("2d");
    ctx.drawImage(imagemFundo, 0, 0);

    // Área central reservada aos dados do usuário.
    const escalaX = canvas.width / 2048;
    const escalaY = canvas.height / 683;

    const x = 995 * escalaX;
    const y = 240 * escalaY;
    const espacamento = 48 * escalaY;

    ctx.fillStyle = "#ffffff";
    ctx.textBaseline = "top";
    ctx.font = `bold ${Math.round(34 * escalaY)}px Arial`;
    ctx.fillText(userInfo.name || "", x, y);

    ctx.font = `${Math.round(29 * escalaY)}px Arial`;
    ctx.fillText(userInfo.department || "", x, y + espacamento);

    ctx.fillText(
        userInfo.job || "",
        x,
        y + espacamento * 2
    );

    ctx.fillText(
        userInfo.email || "",
        x + 70 * escalaX,
        y + espacamento * 3
    );

    return canvas.toDataURL("image/png");
}
