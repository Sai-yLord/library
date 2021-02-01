from flask import Flask, render_template, request
from openpyxl import load_workbook

app = Flask(__name__)

@app.route("/")
def homepage():
    return render_template("index.html")


@app.route("/books/")
def books():
    excel = load_workbook("tales1.xlsx")
    page = excel["Лист1"]

    # tales = [tale.value for tale in page["A"]][1:]
    tales = []
    for tale in page["A"][1:]:
        tales.append(tale.value)

    authors = [author.value for author in page["B"]][1:]   

    html = """
        <a href="/authors">Авторы</a>
        <a href="/books">Книги</a>
        <h1 style="color: red">Тут будет список книг:</h1>
    """

    for i in range(len(tales)):
        html += f"<h2>{tales[i]} - {authors[i]}</h2>"

    return html


@app.route("/authors/")
def authors():
    excel = load_workbook("tales1.xlsx")
    page = excel["Лист1"]
    authors = {author.value for author in page["B"][1:]}
    return render_template("authors.html", authors=authors)

@app.route("/add/", methods=["POST"])
def add():
    f = request.form
    print(f["author"], f["book"])
    return "форма получена"