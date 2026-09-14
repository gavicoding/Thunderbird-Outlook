"""Gera o ícone do aplicativo: um envelope com uma seta de migração,
nas cores do Thunderbird (laranja) fundindo pro Outlook (azul).
Com gradiente diagonal, sombra suave sob o envelope e um brilho glacê
no topo, pra não ficar chapado na barra de tarefas."""
from PIL import Image, ImageDraw, ImageFilter, ImageOps, ImageChops

AZUL_CLARO = (40, 145, 223)
AZUL_ESCURO = (0, 88, 165)
LARANJA_CLARO = (255, 170, 64)
LARANJA_ESCURO = (230, 114, 0)
BRANCO = (255, 255, 255)
CINZA_ENVELOPE = (211, 225, 239)


def gradiente_diagonal(tamanho, cor_ini, cor_fim):
    """Gradiente a 45°, de cor_ini (canto superior esquerdo) a cor_fim (inferior direito)."""
    base = Image.linear_gradient("L").resize((tamanho * 2, tamanho * 2))
    base = base.rotate(45, resample=Image.BICUBIC)
    largura, altura = base.size
    esquerda, topo = (largura - tamanho) // 2, (altura - tamanho) // 2
    base = base.crop((esquerda, topo, esquerda + tamanho, topo + tamanho))
    return ImageOps.colorize(base, black=cor_ini, white=cor_fim).convert("RGBA")


def desenhar_base(tamanho):
    escala = 8
    s = tamanho * escala
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    preto = Image.new("RGBA", (s, s), (0, 0, 0, 255))
    branco_solido = Image.new("RGBA", (s, s), (255, 255, 255, 255))

    margem = int(s * 0.06)
    raio = int(s * 0.20)

    mascara_fundo = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mascara_fundo).rounded_rectangle(
        [margem, margem, s - margem, s - margem], radius=raio, fill=255
    )
    fundo = gradiente_diagonal(s, AZUL_CLARO, AZUL_ESCURO)
    img.paste(fundo, (0, 0), mascara_fundo)

    # brilho glacê suave perto do topo, tipo botão de vidro
    brilho = Image.new("L", (s, s), 0)
    ImageDraw.Draw(brilho).ellipse(
        [int(s * 0.05), int(-s * 0.35), int(s * 0.95), int(s * 0.38)], fill=255
    )
    brilho = brilho.filter(ImageFilter.GaussianBlur(s * 0.03))
    brilho = ImageChops.multiply(brilho, mascara_fundo).point(lambda p: int(p * 0.16))
    img = Image.composite(branco_solido, img, brilho)

    # envelope
    env_margem_x = int(s * 0.17)
    env_topo = int(s * 0.31)
    env_base = int(s * 0.73)
    raio_env = int(s * 0.035)

    # sombra suave projetada pelo envelope sobre o fundo
    sombra = Image.new("L", (s, s), 0)
    deslocamento = int(s * 0.022)
    ImageDraw.Draw(sombra).rounded_rectangle(
        [env_margem_x, env_topo + deslocamento, s - env_margem_x, env_base + deslocamento],
        radius=raio_env, fill=130,
    )
    sombra = sombra.filter(ImageFilter.GaussianBlur(s * 0.025))
    sombra = ImageChops.multiply(sombra, mascara_fundo)
    img = Image.composite(preto, img, sombra)

    # corpo do envelope, com leve gradiente (branco -> cinza-azulado bem claro)
    mascara_env = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mascara_env).rounded_rectangle(
        [env_margem_x, env_topo, s - env_margem_x, env_base], radius=raio_env, fill=255
    )
    corpo_env = gradiente_diagonal(s, BRANCO, CINZA_ENVELOPE)
    img.paste(corpo_env, (0, 0), mascara_env)

    # aba do envelope, com gradiente laranja (remete ao Thunderbird)
    pontos_aba = [
        (env_margem_x, env_topo),
        (s // 2, int(s * 0.50)),
        (s - env_margem_x, env_topo),
    ]
    mascara_aba = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mascara_aba).polygon(pontos_aba, fill=255)
    aba = gradiente_diagonal(s, LARANJA_CLARO, LARANJA_ESCURO)
    img.paste(aba, (0, 0), mascara_aba)

    # sombra sutil de dobra, onde a aba encontra o envelope
    dobra = Image.new("L", (s, s), 0)
    ImageDraw.Draw(dobra).line(pontos_aba, fill=90, width=max(2, int(s * 0.012)))
    dobra = dobra.filter(ImageFilter.GaussianBlur(s * 0.008))
    dobra = ImageChops.multiply(dobra, mascara_env)
    img = Image.composite(preto, img, dobra)

    # seta de migração, com pontas arredondadas e leve sombra
    seta_y = int(s * 0.855)
    seta_x_ini = int(s * 0.29)
    seta_x_fim = int(s * 0.71)
    espessura = max(2, int(s * 0.05))
    ponta = int(s * 0.095)

    def desenhar_seta(desenho, cor, y_desloc=0):
        y = seta_y + y_desloc
        desenho.line([(seta_x_ini, y), (seta_x_fim, y)], fill=cor, width=espessura)
        r = espessura // 2
        desenho.ellipse([seta_x_ini - r, y - r, seta_x_ini + r, y + r], fill=cor)
        desenho.polygon(
            [
                (seta_x_fim + ponta, y),
                (seta_x_fim - int(ponta * 0.1), y - ponta),
                (seta_x_fim - int(ponta * 0.1), y + ponta),
            ],
            fill=cor,
        )

    sombra_seta = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    desenhar_seta(ImageDraw.Draw(sombra_seta), (0, 0, 0, 100), y_desloc=int(s * 0.014))
    sombra_seta = sombra_seta.filter(ImageFilter.GaussianBlur(s * 0.012))
    img = Image.alpha_composite(img, sombra_seta)

    camada_seta = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    desenhar_seta(ImageDraw.Draw(camada_seta), BRANCO + (255,))
    img = Image.alpha_composite(img, camada_seta)

    return img.resize((tamanho, tamanho), Image.LANCZOS)


tamanhos = [16, 24, 32, 48, 64, 128, 256]
imagens = [desenhar_base(t) for t in tamanhos]
imagens[-1].save(
    "icone.ico",
    format="ICO",
    sizes=[(t, t) for t in tamanhos],
)
imagens[-1].save("icone.png")
print("Ícone gerado: icone.ico / icone.png")
