import pathlib
import dash
import dash_core_components as dcc
import dash_html_components as html
import pandas as pd
import plotly.graph_objs as go
import datetime as dt
from dash.dependencies import Output, Input


external_css = "https://codepen.io/chriddyp/pen/bWLwgP.css"


app = dash.Dash(__name__,csrf_protect=False)
app.css.append_css({"external_url": external_css})
app.config.supress_callback_exceptions = True
server = app.server

colors = {"background": "#aa84d1",'plot_bg':'#959c96', "text": "#2a2e2b",'heading':'#d62728'}
text_font = {"fontSize": 20}

PATH = pathlib.Path(__file__).parent
DATA_PATH = PATH.joinpath("data").resolve()

df1 = pd.read_excel(
    DATA_PATH.joinpath("P11-MegaMerchandise.xlsx"), sheet_name="ListOfOrders"
)
df2 = pd.read_excel(
    DATA_PATH.joinpath("P11-MegaMerchandise.xlsx"), sheet_name="OrderBreakdown"
)
df = pd.merge(df1, df2, how="left", left_on="Order_ID", right_on="Order_ID")

Category_options = [{"label": i, "value": i} for i in df["Category"].unique()]

Furniture = df[df["Category"] == "Furniture"]["Sub_Category"].unique().tolist()
Office_Supplies = (
    df[df["Category"] == "Office Supplies"]["Sub_Category"].unique().tolist()
)
Technology = df[df["Category"] == "Technology"]["Sub_Category"].unique().tolist()

cat_sub_cat = {
    "Furniture": Furniture,
    "Office Supplies": Office_Supplies,
    "Technology": Technology,
}


app.layout = html.Div(
    style={"backgroundColor": colors["background"],'marginRight':200},
    children=[
        html.Div(
            [
                html.H1(id='heading',className="container",children="Marketing KPI Dashboard",style={"textAlign": "center","color": colors["text"],"fontSize": text_font})
            ]
        ),
        html.Div(
            [
                html.Div([
                        html.P(children="Product Category",style={"textAlign": "center","color": colors["text"],"fontSize": text_font}),
                        dcc.Dropdown(id="Category",options=Category_options,value="Office Supplies")],className="five columns "),
                html.Div([html.P(id='dcc_subcat',children="Product Sub_Category",style={"textAlign": "center","color": colors["text"],"fontSize": text_font}),
                        dcc.Dropdown(id="Sub-Category")],className="five columns"),
            ],id='dropdown',className='row',style={"marginLeft": 100}),
        html.Div(
            [
                html.Div([html.P(children="Sales-Quantity",style={"textAlign": "center","color": colors["text"],"fontSize": text_font}),
                        html.H4(id="Sales-Quantity",style={"textAlign": "center","color": colors["text"],"fontSize": 30})],
                        className="container three columns"),
                html.Div([
                        html.P(children="Sales-Amount",style={"textAlign": "center","color": colors["text"],"fontSize": text_font}),
                        html.H4(id="Sales-Amount",style={"textAlign":"center","color": colors["text"],"fontSize":30})],
                        id='sales_container',className="container three columns"),
                html.Div([
                        html.P(children="Profit",style={"textAlign": "center","color": colors["text"],"fontSize": text_font}),
                        html.H4(id="Profit",style={"textAlign": "center","color": colors["text"],"fontSize": 30}),],
                        id='profit_container',className="container three columns")
            ],style={"marginLeft": 200, "marginTop": 20}),
        html.Div(
            [
                html.Div([
                        html.P(children="Sales By Order Date",style={"textAlign": "center","color": colors["heading"],"fontSize": text_font}),
                        dcc.Graph(id="Sales-by-orderdate")],className="five columns"),
                html.Div([
                        html.P(children="Profit By Order Date",style={"textAlign": "center","color": colors["heading"],"fontSize": text_font}),
                        dcc.Graph(id="Profit-by-orderdate")],className="five columns"),
            ],id='First_Graph_Block',className='row',style={'marginLeft':100}),
        html.Div(
            [
                html.Div([
                        html.P(children="Order By Country",style={"textAlign": "center","color": colors["heading"],"fontSize": text_font}),
                        dcc.Graph(id="Sales-by-profit")],className="five columns"),
                html.Div([html.P(children="Shipping Mode Distribution",style={"textAlign": "center","color": colors["heading"],"fontSize": text_font}),
                        dcc.Graph(id="Quantity-by-shipmode")],className="five columns"),
            ],id='Second_Graph_Block',className='row',style={'marginLeft':100})
    ],
    className="twelve columns"
)


@app.callback(Output("Sub-Category", "options"), [Input("Category", "value")])
def set_sub_category_options(selected_category):
    return [{"label": i, "value": i} for i in cat_sub_cat[selected_category]]


@app.callback(Output("Sub-Category", "value"), [Input("Sub-Category", "options")])
def set_sub_category_values(sub_category):
    return sub_category[0]["value"]


@app.callback(Output("Sales-Quantity", "children"), [Input("Sub-Category", "value")])
def sales_quantity(sub_category):
    return df[df["Sub_Category"] == sub_category]["Quantity"].sum()


@app.callback(Output("Sales-Amount", "children"), [Input("Sub-Category", "value")])
def sales_amount(sub_category):
    return df[df["Sub_Category"] == sub_category]["Sales"].sum()


@app.callback(Output("Profit", "children"), [Input("Sub-Category", "value")])
def sales_amount(sub_category):
    return df[df["Sub_Category"] == sub_category]["Profit"].sum()


@app.callback(Output("Sales-by-orderdate", "figure"), [Input("Sub-Category", "value")])
def update_figure(sub_category):
    df["Date"] = df["Order_Date"].map(lambda x: x.strftime("%b-%y"))
    filtered_df = df[df["Sub_Category"] == sub_category]
    traces = []
    for i in filtered_df.Sub_Category.unique():
        traces.append(
            go.Scatter(
                x=filtered_df.groupby("Date")["Sales"].sum().index,
                y=filtered_df.groupby("Date")["Sales"].sum().values,
                mode="markers+lines",
            )
        )
        return {
            "data": traces,
            "layout": {
                "margin": {"l": 50, "r": 50, "t": 20, "b": 80},
                "plot_bgcolor": colors["plot_bg"],
                "paper_bgcolor": colors["background"],
                "xaxis": {"title": "Order Date"},
                "yaxis": {"title": "Sales"},
            },
        }


@app.callback(Output("Profit-by-orderdate", "figure"), [Input("Sub-Category", "value")])
def update_figure(sub_category):
    df["Date"] = df["Order_Date"].map(lambda x: x.strftime("%b-%y"))
    filtered_df = df[df["Sub_Category"] == sub_category]
    traces = []
    for i in filtered_df.Sub_Category.unique():
        traces.append(
            go.Scatter(
                x=filtered_df.groupby("Date")["Profit"].sum().index,
                y=filtered_df.groupby("Date")["Profit"].sum().values,
                mode="markers+lines",
            )
        )
        return {
            "data": traces,
            "layout": {
                "margin": {"l": 50, "r": 50, "t": 20, "b": 80},
                "plot_bgcolor": colors["plot_bg"],
                "paper_bgcolor": colors["background"],
                "xaxis": {"title": "Order Date"},
                "yaxis": {"title": "Profit"},
            },
        }


@app.callback(Output("Sales-by-profit", "figure"), [Input("Sub-Category", "value")])
def update_figure(sub_category):
    df["Date"] = df["Order_Date"].map(lambda x: x.strftime("%b-%y"))
    filtered_df = df[df["Sub_Category"] == sub_category]
    traces = []
    for i in filtered_df.Sub_Category.unique():
        traces.append(
            go.Scatter(
                x=filtered_df.groupby("Date")["Profit"].sum().values,
                y=filtered_df.groupby("Date")["Sales"].sum().values,
                mode="markers",
                marker=dict(size=20, opacity=0.5),
            )
        )
        return {
            "data": traces,
            "layout": {
                "margin": {"l": 50, "r": 50, "t": 20, "b": 50},
                "plot_bgcolor": colors["plot_bg"],
                "paper_bgcolor": colors["background"],
                "xaxis": {"title": "Profit"},
                "yaxis": {"title": "Sales"},
            },
        }


@app.callback(
    Output("Quantity-by-shipmode", "figure"), [Input("Sub-Category", "value")]
)
def update_figure(sub_category):
    filtered_df = df[df["Sub_Category"] == sub_category]
    ship_mode = filtered_df.groupby("Ship_Mode")["Quantity"].sum()
    labels = ship_mode.index
    values = ship_mode.values
    fig = {
        "data": [{"values": values, "labels": labels, "hole": 0.5, "type": "pie"}],
        "layout": {
            "plot_bgcolor": colors["plot_bg"],
            "paper_bgcolor": colors["plot_bg"],
            "margin": {"l": 50, "r": 50, "t": 20, "b": 50},
        },
    }
    return fig


if __name__ == "__main__":
    app.run_server(debug=True, threaded=True)
