from flask import Flask, request, jsonify
import win32print
import socket

app = Flask(__name__)

@app.route('/imprimirEtiqueta', methods=['POST'])
def imprimir_etiqueta():
    dados = request.get_json()
    nome = dados.get('nome')
    endereco = dados.get('endereco')
    cidade = dados.get('cidade')
    estado = dados.get('estado')
    cep = dados.get('cep')
    impressora = dados.get('impressora')

    etiqueta_texto = f"""
    <html>
        <header>
        <style>
            
        </style>
        </header>
        <body>
            <p class='nome'><strong>Nome:</strong> {nome}</p>
            <p class='endereco'><strong>Endereço:</strong> {endereco}</p>
            <p class='cidade'><strong>Cidade:</strong> {cidade}</p>
            <p class='estado'><strong>Estado:</strong> {estado}</p>
            <p class='cep'><strong>CEP:</strong> {cep}</p>
        </body>
    </html>
    """

    try:
        hPrinter = win32print.OpenPrinter(impressora)
        try:
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta", None, "RAW"))
            try:
                win32print.StartPagePrinter(hPrinter)
                win32print.WritePrinter(hPrinter, etiqueta_texto.encode('utf-8'))
                win32print.EndPagePrinter(hPrinter)
            finally:
                win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)

        return jsonify({"status": "sucesso", "mensagem": "Etiqueta enviada para impressão."})
    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e)}), 500

if __name__ == '__main__':
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    app.run(host=local_ip, port=5544)
