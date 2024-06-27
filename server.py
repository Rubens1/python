from flask import Flask, request, jsonify
import win32print

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
    Nome: {nome}
    Endereço: {endereco}
    Cidade: {cidade}
    Estado: {estado}
    CEP: {cep}
    """

    try:
        #printer_name = win32print.GetDefaultPrinter()

        # Abrir a impressora
        hPrinter = win32print.OpenPrinter(impressora)
        try:
            # Iniciar um novo trabalho de impressão
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta", None, "RAW"))
            try:
                # Iniciar a página de impressão
                win32print.StartPagePrinter(hPrinter)

                # Escrever na impressora
                win32print.WritePrinter(hPrinter, etiqueta_texto.encode())

                # Finalizar a página de impressão
                win32print.EndPagePrinter(hPrinter)
            finally:
                # Finalizar o trabalho de impressão
                win32print.EndDocPrinter(hPrinter)
        finally:
            # Fechar a impressora
            win32print.ClosePrinter(hPrinter)

        return jsonify({"status": "sucesso", "mensagem": "Etiqueta enviada para impressão."})
    except Exception as e:
        return jsonify({"status": "erro", "mensagem": str(e)}), 500

if __name__ == '__main__':
    app.run(port=3000)
