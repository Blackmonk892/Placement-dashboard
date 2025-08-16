import dash
from dash import dcc, html, Input, Output
import dash_bootstrap_components as dbc
import base64
from io import BytesIO
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import panel as pn

pn.extension('plotly')

# Load data
df = pd.read_excel('placement.xlsx', sheet_name='Sheet1')
df = df.drop(columns=['Details', '#', 'Additional Info'])
df.fillna({'Role': 'NAN', 'CTC': 'NAN', 'Company': 'NAN', 'Stpd': 'NAN'}, inplace=True)
df['CGPA'] = pd.to_numeric(df['CGPA'], errors='coerce')
df['CTC'] = pd.to_numeric(df['CTC'], errors='coerce')
df['Stpd'] = pd.to_numeric(df['Stpd'], errors='coerce')

# Helper to convert matplotlib figure to base64 image
def fig_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    buf.close()
    plt.close(fig)
    return f"data:image/png;base64,{img_base64}"

# Plot Wrapper
def make_plot_block(title, func):
    return dbc.Card([
        dbc.CardHeader(html.H5(title)),
        dbc.CardBody(html.Img(src=fig_to_base64(func().object), style={"width": "100%", "maxWidth": "1000px"}))
    ], className="mb-4")

# Import all plot functions from previous code (assume already defined)
# e.g., plot_hist_by_branch, plot_hist_by_gender, etc.  # Make sure your plot functions are in plots.py

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import panel as pn
from statsmodels.graphics.mosaicplot import mosaic

pn.extension('plotly')

# Load data
df = pd.read_excel('placement.xlsx', sheet_name='Sheet1')
df = df.drop(columns=['Details', '#', 'Additional Info'])

df['Role'] = df['Role'].fillna('NAN')
df['CTC'] = df['CTC'].fillna('NAN')
df['Company'] = df['Company'].fillna('NAN')
df['Stpd'] = df['Stpd'].fillna('NAN')

df['CGPA'] = pd.to_numeric(df['CGPA'], errors='coerce')
df['CTC'] = pd.to_numeric(df['CTC'], errors='coerce')
df['Stpd'] = pd.to_numeric(df['Stpd'], errors='coerce')

# Define plot functions
def plot_hist_by_branch():
    plt.figure(figsize=(10,6))
    sns.histplot(data=df, x='CGPA', hue='Branch', kde=True)
    plt.xlabel('CGPA')
    plt.ylabel('Frequency')
    plt.title('CGPA Distribution by Branches')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_hist_by_gender():
    plt.figure(figsize=(10,6))
    sns.histplot(data=df, x='CGPA', hue='Gender', kde=True)
    plt.xlabel('CGPA')
    plt.ylabel('Frequency')
    plt.title('CGPA Distribution by Gender')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_box_by_role():
    plt.figure(figsize=(12,6))
    sns.boxplot(x='Role', y='CGPA', data=df, palette='Set2')
    plt.xticks(rotation=45)
    plt.title('CGPA distribution by Role')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_box_by_branch():
    plt.figure(figsize=(12,6))
    sns.boxplot(x='Branch', y='CGPA', data=df, palette='Set2')
    plt.xticks(rotation=45)
    plt.title('CGPA distribution by Branch')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_box_by_role():
    plt.Figure(figsize=(12,6))
    sns.boxplot(x='Role', y='CGPA', data= df, palette='Set2')
    plt.xlabel('Placement Role')
    plt.ylabel('CGPA')
    plt.xticks(rotation = 45)
    plt.title('CGPA distribution by Role')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

    

def plot_scatter_gender():
    plt.figure(figsize=(10, 6))
    sns.scatterplot(x='CGPA', y='CTC', data=df, hue='Gender', palette='coolwarm')
    plt.title('CGPA vs. CTC (Gender)')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_scatter_branch():
    plt.figure(figsize=(12, 8))
    sns.scatterplot(x='CGPA', y='CTC', data=df, hue='Branch', size='CTC', style='Branch', alpha=0.6)
    plt.title('CGPA vs. CTC (Branch)')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_hex():
    plt.figure()
    sns.jointplot(data=df, x='CGPA', y='CTC', kind='hex', height=8, gridsize=50, cmap="Blues")
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_count_branch():
    plt.figure(figsize=(12,6))
    sns.countplot(x='Branch', data=df, palette='Set2')
    plt.xticks(rotation=45)
    plt.title('Placement Count by Branch')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_count_branch_gender():
    plt.figure(figsize=(12,6))
    sns.countplot(x='Branch', data=df, hue='Gender')
    plt.xticks(rotation=45)
    plt.title('Placement Count by Branch and Gender')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_count_gender():
    plt.figure(figsize=(8,5))
    sns.countplot(x='Gender', data=df, hue='Branch', palette='Set1')
    plt.title('Placement Count by Gender')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_count_role():
    plt.Figure(figsize=(14,6))
    sns.countplot(x='Role', data=df, palette='viridis')
    plt.xlabel('Placement Role')
    plt.ylabel('Number od Students')
    plt.xticks(rotation = 45)
    plt.title('placement count by Role')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)


def plot_mosaic():
    mosaic_data = df.groupby(['Branch','Role']).size().unstack(fill_value=0)
    mosaic_dict = {(branch, role): count for branch, row in mosaic_data.iterrows() for role, count in row.items()}
    plt.figure()
    mosaic(mosaic_dict, title='Mosaic Plot of placement roles by branch')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_avg_ctc_branch():
    avg_ctc = df.groupby('Branch')['CTC'].mean().reset_index()
    plt.figure(figsize=(12,6))
    sns.barplot(x='Branch', y='CTC', data=avg_ctc, palette='Set2')
    plt.xticks(rotation=45)
    plt.title('Average CTC per Branch')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_median_ctc_branch():
    median_ctc = df.groupby('Branch')['CTC'].median().reset_index()
    plt.figure(figsize=(12,6))
    sns.barplot(x='Branch', y='CTC', data=median_ctc, palette='Set2')
    plt.xticks(rotation=45)
    plt.title('Median CTC per Branch')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_avg_ctc_role():
    role_ctc = df.groupby('Role')['CTC'].mean().reset_index()
    plt.figure(figsize=(12,6))
    sns.barplot(x='Role', y='CTC', data=role_ctc, palette = 'Set2')
    plt.title('Average CTC on the basis of role')
    plt.xlabel('Role')
    plt.ylabel('Average CTC (LPA)')
    plt.xticks(rotation=45)
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_median_ctc_role():
    role_ctc = df.groupby('Role')['CTC'].median().reset_index()
    plt.figure(figsize=(12,6))
    sns.barplot(x='Role', y='CTC', data=role_ctc, palette = 'Set2')
    plt.title('Median CTC on the basis of role')
    plt.xlabel('Role')
    plt.ylabel('Median CTC (LPA)')
    plt.xticks(rotation=45)
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_avg_ctc_role_stpd():
    role_ctc = df.groupby('Role')['Stpd'].mean().reset_index()
    plt.figure(figsize=(12,6))
    sns.barplot(x='Role', y='Stpd', data=role_ctc, palette = 'Set2')
    plt.title('Average stipend on the basis of role')
    plt.xlabel('Role')
    plt.ylabel('Average stipend (K)')
    plt.xticks(rotation=45)
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_median_ctc_role_stpd():
    role_ctc = df.groupby('Role')['Stpd'].median().reset_index()
    plt.figure(figsize=(12,6))
    sns.barplot(x='Role', y='Stpd', data=role_ctc, palette = 'Set2')
    plt.title('median stipend on the basis of role')
    plt.xlabel('Role')
    plt.ylabel('median stipend (K)')
    plt.xticks(rotation=45)
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_median_ctc_role_gender():
    role_ctc = df.groupby('Gender')['CTC'].median().reset_index()
    plt.figure(figsize=(12,6))
    sns.barplot(x='Gender', y='CTC', data=role_ctc, palette = 'Set2')
    plt.title('Median CTC vs Gender')
    plt.xlabel('Role')
    plt.ylabel('Median CTC (LPA)')
    plt.xticks(rotation=45)
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_violin_role():
    plt.figure()
    sns.violinplot(x='Role', y='CTC', data=df, inner='quartile')
    plt.xticks(rotation=45)
    plt.title("CTC Distribution by Role")
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_corr():
    plt.figure(figsize=(8,6))
    sns.heatmap(df[['CGPA', 'CTC', 'Stpd']].corr(), annot=True, cmap='coolwarm', fmt=".2f")
    plt.title('Correlation Heatmap: CGPA, CTC, Stpd')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_pie_branch():
    plt.figure()
    df['Branch'].value_counts().plot.pie(autopct='%1.1f%%', startangle=90, cmap='coolwarm')
    plt.title('Placement Distribution by Branch')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_pie_gender():
    plt.figure()
    df['Gender'].value_counts().plot.pie(autopct='%1.1f%%', startangle=90, colors=['pink', 'blue'])
    plt.title('Placement Distribution by Gender')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_pie_role():
    role_counts = df['Role'].value_counts().nlargest(6)
    plt.Figure(figsize=(10,6))
    role_counts.plot.pie(autopct='%1.1f%%', startangle=90, cmap='coolwarm')
    plt.title('Placement Distribution by Role(largest 6)')
    plt.ylabel('')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_pair_branch():
    sns.pairplot(df, vars=['CGPA','CTC','Stpd'], hue='Branch', palette='Set1')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_pair_gender():
    sns.pairplot(df, vars=['CGPA','CTC','Stpd'], hue='Gender', palette='Set1')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

# App Init with Bootstrap
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
app.title = "DTU Placement Dashboard"

# Layout
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col(html.H1("DTU Placement Dashboard", className="text-center text-primary mb-4"), width=12)
    ]),
    dbc.Row([
        dbc.Col([
            dcc.Tabs([
                dcc.Tab(label='Distribution', children=[
                    make_plot_block("CGPA by Branch", plot_hist_by_branch),
                    make_plot_block("CGPA by Gender", plot_hist_by_gender),
                ]),
                dcc.Tab(label='Boxplots', children=[
                    make_plot_block("CGPA by Branch", plot_box_by_branch),
                    make_plot_block("CGPA by Role", plot_box_by_role),
                ]),
                dcc.Tab(label='Scatter & Hex', children=[
                    make_plot_block("CTC vs CGPA by Gender", plot_scatter_gender),
                    make_plot_block("CTC vs CGPA by Branch", plot_scatter_branch),
                    make_plot_block("Hex Plot", plot_hex),
                ]),
                dcc.Tab(label='Counts', children=[
                    make_plot_block("Placements by Branch", plot_count_branch),
                    make_plot_block("Placements by Branch and Gender", plot_count_branch_gender),
                    make_plot_block("Placements by Gender", plot_count_gender),
                    make_plot_block("Placements by Role", plot_count_role),
                ]),
                dcc.Tab(label='CTC & Stipend Analysis', children=[
                    make_plot_block("Avg CTC by Branch", plot_avg_ctc_branch),
                    make_plot_block("Median CTC by Branch", plot_median_ctc_branch),
                    make_plot_block("Avg CTC by Role", plot_avg_ctc_role),
                    make_plot_block("Median CTC by Role", plot_median_ctc_role),
                    make_plot_block("Avg Stipend by Role", plot_avg_ctc_role_stpd),
                    make_plot_block("Median Stipend by Role", plot_median_ctc_role_stpd),
                    make_plot_block("Median CTC by Gender", plot_median_ctc_role_gender),
                    make_plot_block("Violin Plot by Role", plot_violin_role),
                ]),
                dcc.Tab(label='Pie Charts', children=[
                    make_plot_block("Branch Distribution", plot_pie_branch),
                    make_plot_block("Gender Distribution", plot_pie_gender),
                    make_plot_block("Top 6 Roles Distribution", plot_pie_role),
                ]),
                dcc.Tab(label='Pair & Correlation', children=[
                    make_plot_block("Correlation Heatmap", plot_corr),
                    make_plot_block("Pairplot by Branch", plot_pair_branch),
                    make_plot_block("Pairplot by Gender", plot_pair_gender),
                ]),
                dcc.Tab(label='Mosaic', children=[
                    make_plot_block("Mosaic Plot of Roles by Branch", plot_mosaic),
                ]),
            ])
        ], width=12)
    ])
], fluid=True)

if __name__ == '__main__':
    app.run(debug=True)
