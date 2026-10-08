
from flask import Flask, request, render_template_string, redirect, send_file
from datetime import datetime, timedelta
import sqlite3
import pandas as pd
import os

app = Flask(__name__)

DB = "os.db"


# =========================
# BANCO DE DADOS
# =========================

def init_db():
    conn = sqlite3.connect(DB)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS os_registradas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT,
            data TEXT,
            inicio TEXT,
            fim TEXT,
            responsavel TEXT,
            horas TEXT,
            segundos REAL
        )
    """)

    conn.commit()
    conn.close()


# =========================
# FORMATAÇÃO DE HORAS
# =========================

def formatar_horas(segundos):
    segundos = int(segundos)

    horas = segundos // 3600
    minutos = (segundos % 3600) // 60

    return f"{horas}h {minutos:02d}min"


# Inicializa o banco
init_db()


# =========================
# HTML
# =========================

HTML = """
<!DOCTYPE html>
<html lang="pt-BR">

<head>

<meta charset="UTF-8">

<meta
    name="viewport"
    content="width=device-width, initial-scale=1.0"
>

<title>Central de OS</title>


<style>

/* =========================
   CONFIGURAÇÕES GERAIS
   ========================= */

* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
    font-family: Arial, sans-serif;
}

body {
    background: #ED6B51;
    padding: 10px;
    overflow-x: hidden;
}

.container {
    width: 100%;
    max-width: 1200px;
    background: white;
    margin: auto;
    padding: 20px;
    border-radius: 20px;
}


/* =========================
   TÍTULOS
   ========================= */

h1 {
    text-align: center;
    color: #F54927;
    margin-bottom: 25px;
    font-size: 30px;
}

h2 {
    text-align: center;
    color: #F54927;
    margin-top: 30px;
    margin-bottom: 15px;
}


/* =========================
   FORMULÁRIO
   ========================= */

form {
    background: #f7f7f7;
    padding: 18px;
    border-radius: 15px;
}

label {
    display: block;
    margin-bottom: 6px;
    font-weight: bold;
    color: #555;
}

input {
    width: 100%;
    padding: 14px;
    margin-bottom: 14px;

    border: 1px solid #ccc;
    border-radius: 10px;

    font-size: 16px;
    outline: none;
}

input:focus {
    border-color: #F54927;
}


/* =========================
   BOTÃO REGISTRAR
   ========================= */

.btn-registrar {
    background: #F54927;
    color: white;

    border: none;
    padding: 15px;

    width: 100%;

    font-size: 18px;
    font-weight: bold;

    border-radius: 10px;

    cursor: pointer;
}

.btn-registrar:hover {
    background: #d63f21;
}


/* =========================
   CARDS
   ========================= */

.cards {
    display: flex;
    gap: 15px;

    margin-top: 20px;
    margin-bottom: 20px;
}

.card {
    flex: 1;

    background: #fff4f1;

    border-left: 5px solid #F54927;

    border-radius: 10px;

    padding: 15px;

    text-align: center;
}

.card h3 {
    color: #666;
    font-size: 14px;
}

.card p {
    color: #F54927;

    font-size: 24px;
    font-weight: bold;

    margin-top: 5px;
}


/* =========================
   TABELAS
   ========================= */

.table-container {
    width: 100%;
    overflow-x: auto;
}

table {
    width: 100%;

    border-collapse: collapse;

    margin-top: 15px;

    min-width: 600px;
}

th {
    background: #F54927;
    color: white;
}

th,
td {
    border: 1px solid #ccc;

    padding: 10px;

    text-align: center;
}

tr:nth-child(even) {
    background: #f5f5f5;
}


/* =========================
   BOTÃO EXCEL
   ========================= */

.exportar-container {
    text-align: center;
    margin: 25px 0;
}

.btn-excel {
    display: inline-block;

    background: #28A745;

    color: white;

    text-decoration: none;

    border: none;

    padding: 13px 25px;

    border-radius: 8px;

    cursor: pointer;

    font-size: 16px;

    font-weight: bold;

    transition: 0.2s;
}

.btn-excel:hover {
    background: #218838;

    transform: scale(1.03);
}


/* =========================
   BOTÃO EXCLUIR
   ========================= */

.btn-excluir {
    display: inline-block;

    background: #dc3545;

    color: white;

    text-decoration: none;

    border: none;

    padding: 8px 12px;

    border-radius: 6px;

    cursor: pointer;

    font-weight: bold;

    font-size: 13px;
}

.btn-excluir:hover {
    background: #b02a37;
}


/* =========================
   RESPONSIVIDADE
   ========================= */

@media (max-width: 768px) {

    body {
        padding: 5px;
    }

    .container {
        padding: 15px;
        border-radius: 15px;
    }

    h1 {
        font-size: 22px;
    }

    h2 {
        font-size: 20px;
    }

    .cards {
        flex-direction: column;
    }

    .card p {
        font-size: 20px;
    }

    table {
        font-size: 12px;
    }

    th,
    td {
        padding: 8px;
    }

    input {
        font-size: 16px;
    }

    .btn-registrar {
        font-size: 16px;
    }

    .btn-excel {
        width: 100%;
    }
}

</style>

</head>


<body>


<div class="container">


<h1>
CENTRAL DE O.S JAMAICA
</h1>


<!-- =========================
     FORMULÁRIO
     ========================= -->

<form method="POST">

<input
    type="text"
    name="codigo"
    placeholder="CÓDIGO DA OS"
    required
>


<label>
📅 DATA DA MANUTENÇÃO
</label>

<input
    type="date"
    name="data"
    required
>


<label>
🕒 HORÁRIO DE INÍCIO
</label>

<input
    type="time"
    name="inicio"
    step="60"
    required
>


<label>
🕓 HORÁRIO DE TÉRMINO
</label>

<input
    type="time"
    name="fim"
    step="60"
    required
>


<input
    type="text"
    name="responsavel"
    placeholder="RESPONSÁVEL PELA MANUTENÇÃO"
    required
>


<button
    type="submit"
    class="btn-registrar"
>
    REGISTRAR OS
</button>


</form>


<!-- =========================
     RESUMO
     ========================= -->

<h2>
Resumo Geral
</h2>


<div class="cards">


<div class="card">

<h3>
Total de OS
</h3>

<p>
{{ total_os }}
</p>

</div>


<div class="card">

<h3>
Total de Horas
</h3>

<p>
{{ total_horas_formatado }}
</p>

</div>


<div class="card">

<h3>
Média por OS
</h3>

<p>
{{ media_formatada }}
</p>

</div>


</div>


<!-- =========================
     HORAS POR RESPONSÁVEL
     ========================= -->

<h2>
Horas por Responsável
</h2>


<div class="table-container">

<table>

<tr>

<th>
Responsável
</th>

<th>
Horas
</th>

</tr>


{% for nome, horas in responsaveis.items() %}

<tr>

<td>
{{ nome }}
</td>

<td>
{{ horas }}
</td>

</tr>

{% endfor %}


</table>

</div>


<!-- =========================
     BOTÃO EXPORTAR EXCEL
     ========================= -->

<div class="exportar-container">

<a
    href="/exportar_excel"
    class="btn-excel"
>
    📊 EXPORTAR PARA EXCEL
</a>

</div>


<!-- =========================
     OS REGISTRADAS
     ========================= -->

<h2>
OS Registradas
</h2>


<div class="table-container">

<table>

<tr>

<th>
Código
</th>

<th>
Data
</th>

<th>
Início
</th>

<th>
Fim
</th>

<th>
Responsável
</th>

<th>
Horas
</th>

<th>
Ações
</th>

</tr>


{% for item in lista %}

<tr>

<td>
{{ item["codigo"] }}
</td>

<td>
{{ item["data"] }}
</td>

<td>
{{ item["inicio"] }}
</td>

<td>
{{ item["fim"] }}
</td>

<td>
{{ item["responsavel"] }}
</td>

<td>
{{ item["horas"] }}
</td>


<td>

<a
    href="/excluir/{{ item['id'] }}"
    class="btn-excluir"
    onclick="return confirm('Deseja realmente excluir esta OS?');"
>
    EXCLUIR
</a>

</td>


</tr>

{% endfor %}


</table>

</div>


</div>


</body>

</html>
"""


# =========================
# PÁGINA PRINCIPAL
# =========================

@app.route("/", methods=["GET", "POST"])
def index():

    # =========================
    # REGISTRAR OS
    # =========================

    if request.method == "POST":

        codigo = request.form["codigo"]

        data = datetime.strptime(
            request.form["data"],
            "%Y-%m-%d"
        ).strftime("%d/%m/%Y")

        inicio = request.form["inicio"]

        fim = request.form["fim"]

        responsavel = request.form["responsavel"]


        # Converte os horários
        hora_inicio = datetime.strptime(
            inicio,
            "%H:%M"
        )

        hora_fim = datetime.strptime(
            fim,
            "%H:%M"
        )


        # Caso termine no dia seguinte
        if hora_fim < hora_inicio:
            hora_fim += timedelta(days=1)


        # Calcula os segundos trabalhados
        segundos = (
            hora_fim - hora_inicio
        ).total_seconds()


        # Formata as horas
        horas_formatadas = formatar_horas(
            segundos
        )


        # Salva no banco
        conn = sqlite3.connect(DB)

        cursor = conn.cursor()


        cursor.execute("""
            INSERT INTO os_registradas
            (
                codigo,
                data,
                inicio,
                fim,
                responsavel,
                horas,
                segundos
            )

            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            codigo,
            data,
            inicio,
            fim,
            responsavel,
            horas_formatadas,
            segundos
        ))


        conn.commit()

        conn.close()


        return redirect("/")


    # =========================
    # BUSCAR OS
    # =========================

    conn = sqlite3.connect(DB)

    conn.row_factory = sqlite3.Row

    cursor = conn.cursor()


    cursor.execute("""
        SELECT *
        FROM os_registradas
        ORDER BY id DESC
    """)


    rows = cursor.fetchall()

    os_registradas = [
        dict(row)
        for row in rows
    ]


    conn.close()


    # =========================
    # TOTAL DE HORAS
    # =========================

    total_segundos = sum(
        float(item["segundos"])
        for item in os_registradas
    )


    total_os = len(
        os_registradas
    )


    # =========================
    # MÉDIA
    # =========================

    if total_os > 0:

        media_segundos = (
            total_segundos / total_os
        )

    else:

        media_segundos = 0


    # =========================
    # HORAS POR RESPONSÁVEL
    # =========================

    horas_responsaveis = {}


    for item in os_registradas:

        nome = item["responsavel"]


        if nome not in horas_responsaveis:

            horas_responsaveis[nome] = 0


        horas_responsaveis[nome] += float(
            item["segundos"]
        )


    responsaveis_formatado = {}


    for nome, segundos in horas_responsaveis.items():

        responsaveis_formatado[nome] = (
            formatar_horas(segundos)
        )


    # =========================
    # RENDERIZA A PÁGINA
    # =========================

    return render_template_string(
        HTML,

        responsaveis=responsaveis_formatado,

        lista=os_registradas,

        total_os=total_os,

        total_horas_formatado=(
            formatar_horas(total_segundos)
        ),

        media_formatada=(
            formatar_horas(media_segundos)
        )
    )


# =========================
# EXCLUIR OS
# =========================

@app.route("/excluir/<int:id>")
def excluir(id):

    conn = sqlite3.connect(DB)

    cursor = conn.cursor()


    cursor.execute(
        """
        DELETE FROM os_registradas
        WHERE id=?
        """,
        (id,)
    )


    conn.commit()

    conn.close()


    return redirect("/")


# =========================
# EXPORTAR PARA EXCEL
# =========================

@app.route("/exportar_excel")
def exportar_excel():

    conn = sqlite3.connect(DB)


    df = pd.read_sql_query(
        """
        SELECT

            codigo AS 'Código',

            data AS 'Data',

            inicio AS 'Início',

            fim AS 'Fim',

            responsavel AS 'Responsável',

            horas AS 'Horas Trabalhadas'

        FROM os_registradas

        ORDER BY id DESC
        """,

        conn
    )


    conn.close()


    # Nome do arquivo
    arquivo = "Relatorio_OS.xlsx"


    # Cria o Excel
    with pd.ExcelWriter(
        arquivo,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="OS"
        )


        # Ajusta largura das colunas
        worksheet = writer.sheets["OS"]


        for coluna in worksheet.columns:

            maior_tamanho = 0

            letra_coluna = coluna[0].column_letter


            for celula in coluna:

                try:

                    tamanho = len(
                        str(celula.value)
                    )

                    if tamanho > maior_tamanho:
                        maior_tamanho = tamanho

                except:
                    pass


            worksheet.column_dimensions[
                letra_coluna
            ].width = maior_tamanho + 3


    # Envia o arquivo para download
    return send_file(
        arquivo,
        as_attachment=True,
        download_name="Relatorio_OS.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    )


# =========================
# INICIAR SERVIDOR
# =========================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )