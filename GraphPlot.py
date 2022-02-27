import plotly.graph_objects as go
import plotly.offline as pyo
import json
from ResultOutput import weird_division, TotalResultCountForm, TotalResultCountType

def FormGraphPlot():
    with open("data/json/config.json", "r") as f:
        config = json.load(f)
        FormPath = config["FormPath"]
        FormPathHTML = config['FormPathHTML']
    with open("data/json/FormData.json", "r") as f:
        form = json.load(f)

    # Form
    categories = config["form"]
    categories = [*categories, categories[0]]

    for i, (name, data) in enumerate(form.items(), start=0):
        if name != "None_None(None)":
            list1 = []
            for j, (forms, value) in enumerate(data.items(), start=0):
                years = str(form[name]["year"])
                if forms != "year":
                    list1.append(weird_division(value, TotalResultCountForm(years, forms))*100)
            # print(name1)
            fig = go.Figure(
                data=[
                    go.Scatterpolar(r=list1, theta=categories, fill='toself', name='FormMarks')
                ],
                layout=go.Layout(
                    title=go.layout.Title(text=name+' Form Score'),
                    polar={'radialaxis': {'visible': True, 'range': [0, 100]}},
                    showlegend=True
                ),
            )
            pyo.plot(fig, filename=FormPathHTML + name + ".html", auto_open=False)
            fig.write_image(FormPath + name + ".png", format="png")
            print("Done " + name)

def TypeGraphPlot():
    with open("data/json/config.json", "r") as f:
        config = json.load(f)
        TypePath = config["TypePath"]
        TypePathHTML = config['TypePathHTML']
    with open("data/json/TypeData.json", "r") as f:
        type = json.load(f)

    # Type
    categories = config["type"]
    categories = [*categories, categories[0]]

    for i, (name, data) in enumerate(type.items(), start=0):
        if name != "None_None(None)":
            list1 = []
            for j, (types, value) in enumerate(data.items(), start=0):
                years = str(type[name]["year"])
                if types != "year" or types != "#N/A":
                    list1.append(weird_division(value, TotalResultCountType(years, types))*100)
            # print(list1)
            fig = go.Figure(
                data=[
                    go.Scatterpolar(r=list1, theta=categories, fill='toself', name='FormMarks')
                ],
                layout=go.Layout(
                    title=go.layout.Title(text=name+' Form Score'),
                    polar={'radialaxis': {'visible': True, 'range': [0, 100]}},
                    showlegend=True
                ),
            )
            pyo.plot(fig, filename=TypePathHTML + name + ".html", auto_open=False)
            fig.write_image(TypePath + name + ".png", format="png")
            print("Done " + name)