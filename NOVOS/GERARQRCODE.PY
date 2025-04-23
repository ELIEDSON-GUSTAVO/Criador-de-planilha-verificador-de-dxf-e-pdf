import qrcode

# Texto que será transformado em QR Code
texto = input("Digite o texto ou link para gerar o QR Code: ")

# Geração do QR Code
qr = qrcode.QRCode(
    version=1,  # controle de tamanho (1 a 40)
    error_correction=qrcode.constants.ERROR_CORRECT_H,  # nível de correção de erro
    box_size=10,  # tamanho de cada "caixa" do QR
    border=4,  # espessura da borda
)
qr.add_data(texto)
qr.make(fit=True)

# Criação da imagem
imagem_qr = qr.make_image(fill_color="black", back_color="white")

# Salvar imagem
imagem_qr.save("qrcode_gerado.png")
print("QR Code gerado e salvo como 'qrcode_gerado.png'")
