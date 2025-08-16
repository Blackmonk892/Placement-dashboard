import dash
from dash import dcc, html
import base64
from io import BytesIO
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

# Your plot functions (unchanged except corrected Figure syntax and one return fix)
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
    sns.boxplot(x='Role', y='CGPA', data= df, palette='Set2')
    plt.xlabel('Placement Role')
    plt.ylabel('CGPA')
    plt.xticks(rotation = 45)
    plt.title('CGPA distribution by Role')
    return pn.pane.Matplotlib(plt.gcf(), tight=True)

def plot_box_by_branch():
    plt.figure(figsize=(12,6))
    sns.boxplot(x='Branch', y='CGPA', data=df, palette='Set2')
    plt.xticks(rotation=45)
    plt.title('CGPA distribution by Branch')
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
    plt.figure(figsize=(14,6))
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
    plt.xlabel('Gender')
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
    plt.figure(figsize=(10,6))
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

# Helper
def fig_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    buf.close()
    plt.close(fig)
    return f"data:image/png;base64,{img_base64}"

def make_plot_block(func):
    return html.Div([
        html.H4(func.__name__.replace('_', ' ').title()),
        html.Img(src=fig_to_base64(func().object), style={"width": "100%", "maxWidth": "1000px", "marginBottom": "30px"})
    ])

# App init
app = dash.Dash(__name__)
app.title = "Placement Dashboard"

app.layout = html.Div([
    html.H1("Placement Visualization Dashboard", style={"textAlign": "center"}),
    dcc.Tabs([
        dcc.Tab(label='Distribution', children=[
            make_plot_block(plot_hist_by_branch),
            make_plot_block(plot_hist_by_gender),
        ]),
        dcc.Tab(label='Boxplots', children=[
            make_plot_block(plot_box_by_branch),
            make_plot_block(plot_box_by_role),
        ]),
        dcc.Tab(label='Scatter & Hex', children=[
            make_plot_block(plot_scatter_gender),
            make_plot_block(plot_scatter_branch),
            make_plot_block(plot_hex),
        ]),
        dcc.Tab(label='Counts', children=[
            make_plot_block(plot_count_branch),
            make_plot_block(plot_count_branch_gender),
            make_plot_block(plot_count_gender),
            make_plot_block(plot_count_role),
        ]),
        dcc.Tab(label='CTC & Stipend Analysis', children=[
            make_plot_block(plot_avg_ctc_branch),
            make_plot_block(plot_median_ctc_branch),
            make_plot_block(plot_avg_ctc_role),
            make_plot_block(plot_median_ctc_role),
            make_plot_block(plot_avg_ctc_role_stpd),
            make_plot_block(plot_median_ctc_role_stpd),
            make_plot_block(plot_median_ctc_role_gender),
            make_plot_block(plot_violin_role),
        ]),
        dcc.Tab(label='Pie Charts', children=[
            make_plot_block(plot_pie_branch),
            make_plot_block(plot_pie_gender),
            make_plot_block(plot_pie_role),
        ]),
        dcc.Tab(label='Pair & Correlation', children=[
            make_plot_block(plot_corr),
            make_plot_block(plot_pair_branch),
            make_plot_block(plot_pair_gender),
        ]),
        dcc.Tab(label='Mosaic', children=[
            make_plot_block(plot_mosaic),
        ]),
    ])
])

if __name__ == '__main__':
    app.run(debug=True)
