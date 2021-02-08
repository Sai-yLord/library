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

    object_list = [[tale.value, tale.offset(column=1).value] for tale in page["A"][1:]]
    return render_template("books.htm", object_list=object_list)


    # tales = [tale.value for tale in page["A"]][1:]
    # tales = []
    # for tale in page["A"][1:]:
    #     tales.append(tale.value)

    # authors = [author.value for author in page["B"]][1:]   

    # html = """
    #     <a href="/authors">Авторы</a>
    #     <a href="/books">Книги</a>
    #     <h1 style="color: red">Тут будет список книг:</h1>
    # """

    # for i in range(len(tales)):
    #     html += f"<h2>{tales[i]} - {authors[i]}</h2>"

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
    excel = load_workbook("tales1.xlsx")
    page = excel["Лист1"]
    last = len(page["A"]) + 1
    page[f"A{last}"] = f["book"]
    page[f"B{last}"] = f["author"]
    excel.save("tales1.xlsx")
    return "форма получена"

@app.route("/book<num>")
def book(num):
    excel = load_workbook("tales1.xlsx")
    page = excel["Лист1"]
    object_list = [[tale.value, tale.offset(column=1).value, tale.offset(column=2).value] for tale in page["A"][1:]]
    obj = object_list[int(num)]
    link_list = [[tale.value, tale.offset(column=1).value, tale.offset(column=2).value] for tale in page["D"][1:]]
    link = link_list[int(num)]
    return render_template("book.htm", obj=obj)
