import pandas as pd

# Lendo a planilha
arquivo_excel = 'PLANILHA_NEW.xlsx'
df = pd.read_excel(arquivo_excel)

# Convertendo o DataFrame para uma tabela HTML sem índices
tabela_html = df.to_html(index=False, escape=False)

# Salvando a tabela em um arquivo HTML com um design mais bonito
with open('index.html', 'w', encoding='utf-8') as arquivo_html:
    arquivo_html.write(f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <link rel="shortcut icon" href="coffe.ico" type="image/x-icon">
        <title>Info Produtos</title>
        <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">
        <style>
            body {{
                font-family: 'Roboto', sans-serif;
                background-color: #a3a4a4;
                color: #333;
                margin: 0;
                padding: 20px;
                display: flex;
                justify-content: center;
                align-items: center;
                min-height: 100vh;
            }}
            h1 {{
                text-align: center;
                color: #2c3e50;
                margin-bottom: 20px;
            }}
            table {{
                width: 80%;
                border-collapse: collapse;
                margin: 0 auto;
                background-color: #fff;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
                border-radius: 10px;
                overflow: hidden;
            }}
            th, td {{
                padding: 15px;
                text-align: left;
                border-bottom: 1px solid #ddd;
            }}
            th {{
                background-color: #27ae60;
                color: white;
                text-transform: uppercase;
                font-size: 14px;
                letter-spacing: 1px;
            }}
            td {{
                font-size: 14px;
                color: #555;
            }}
            tr:nth-child(even) {{
                background-color: #f9f9f9;
            }}
            tr:hover {{
                background-color: #f1f1f1;
            }}
            button {{
                display: block;
                width: 200px;
                margin: 20px auto;
                padding: 10px;
                background-color: #27ae60;
                color: white;
                border: none;
                border-radius: 25px;
                font-size: 16px;
                cursor: pointer;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                transition: background-color 0.3s ease;
            }}
            button:hover {{
                background-color: #219150;
            }}
        </style>
    </head>
    <body>
        <div>
            <h1>Nossa Planilha</h1>
            <table contenteditable="true">
                {tabela_html}
            </table>
            <button onclick="salvarTabela()">Salvar Edições</button>
        </div>

        <script>
            function salvarTabela() {{
                let tabela = document.querySelector('table').outerHTML;
                localStorage.setItem('tabelaEditada', tabela);
                alert('Tabela salva localmente!');
            }}

            function carregarTabela() {{
                let tabelaSalva = localStorage.getItem('tabelaEditada');
                if (tabelaSalva) {{
                    document.querySelector('table').outerHTML = tabelaSalva;
                }}
            }}

            window.onload = carregarTabela;
        </script>
    </body>
    </html>
    """)

print("Página HTML bonita gerada com sucesso!")
