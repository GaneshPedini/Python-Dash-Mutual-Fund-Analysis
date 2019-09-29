import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_bootstrap_components as dbc
import pandas as pd
import numpy as np
import openpyxl
import base64

KotakStandard=pd.read_excel('MyPortfolio.xlsx', sheet_name='KotakStandard')
SBIBlueChip=pd.read_excel('MyPortfolio.xlsx', sheet_name='SBIBlueChip')
RelainceSmallCap=pd.read_excel('MyPortfolio.xlsx', sheet_name='RelainceSmallCap')
BirlaSmallCap=pd.read_excel('MyPortfolio.xlsx', sheet_name='BirlaSmallCap')
BirlaELSS=pd.read_excel('MyPortfolio.xlsx', sheet_name='BirlaELSS')

kotak_max_row=KotakStandard.shape[0]
sbi_max_row=SBIBlueChip.shape[0]
reliance_max_row=RelainceSmallCap.shape[0]
Birla_max_row=BirlaSmallCap.shape[0]
BirlaELSS_max_row=BirlaELSS.shape[0]

external_stylesheets =['https://codepen.io/chriddyp/pen/bWLwgP.css']
app=dash.Dash(__name__,external_stylesheets=external_stylesheets)

#Including Image to Dash
image_filename = 'C:\Python3\Scripts\MyScripts\Project\MFPortfolio\Photo.png'
encoded_image = base64.b64encode(open(image_filename, 'rb').read())

kotak_last_nav=KotakStandard.iloc[[kotak_max_row-1],[1]].values
kotak_prev_nav=KotakStandard.iloc[[kotak_max_row-2],[1]].values
kotakdatediff= round(float(kotak_last_nav - kotak_prev_nav),2)
kotak_latest_date="".join(map(str, KotakStandard.iloc[[kotak_max_row-1],[0]].values))[2:-2]

sbi_last_nav=SBIBlueChip.iloc[[sbi_max_row-1],[1]].values
sbi_prev_nav=SBIBlueChip.iloc[[sbi_max_row-2],[1]].values
sbidatediff= round(float(sbi_last_nav - sbi_prev_nav),2)
sbi_latest_date="".join(map(str, SBIBlueChip.iloc[[sbi_max_row-1],[0]].values))[2:-2]

brl_last_nav=BirlaSmallCap.iloc[[Birla_max_row-1],[1]].values
brl_prev_nav=BirlaSmallCap.iloc[[Birla_max_row-2],[1]].values
brldatediff= round(float(brl_last_nav - brl_prev_nav),2)
brl_latest_date="".join(map(str, BirlaSmallCap.iloc[[Birla_max_row-1],[0]].values))[2:-2]

rlc_last_nav=RelainceSmallCap.iloc[[reliance_max_row-1],[1]].values
rlc_prev_nav=RelainceSmallCap.iloc[[reliance_max_row-2],[1]].values
rlcdatediff= round(float(rlc_last_nav - rlc_prev_nav),2)
rlc_latest_date="".join(map(str, RelainceSmallCap.iloc[[reliance_max_row-1],[0]].values))[2:-2]

elss_last_nav=BirlaELSS.iloc[[BirlaELSS_max_row-1],[1]].values
elss_prev_nav=BirlaELSS.iloc[[BirlaELSS_max_row-2],[1]].values
elssdatediff= round(float(elss_last_nav - elss_prev_nav),2)
elss_latest_date="".join(map(str, BirlaELSS.iloc[[BirlaELSS_max_row-1],[0]].values))[2:-2]

kotak_latest_date='Last Close Date:'+ kotak_latest_date
kotak_latest_date=kotak_latest_date.rjust(4)
kotak_last_nav='NAV on Close: '+ "".join(map(str,kotak_last_nav))
kotak_last_nav=kotak_last_nav.rjust(4)
kotak_change="Change in NAV: " +  "["+str(kotakdatediff)+"]"
kotak_change=kotak_change.rjust(4)

sbi_latest_date='Last Close Date:'+ sbi_latest_date
sbi_last_nav='NAV on Close: '+ "".join(map(str,sbi_last_nav)) 
sbi_change="Change in NAV: " + "["+str(sbidatediff)+"]"

brl_latest_date='Last Close Date:'+ brl_latest_date
brl_last_nav='NAV on Close: '+ "".join(map(str,brl_last_nav))
brl_change="Change in NAV:" + "["+str(brldatediff)+ "]"

rlc_latest_date='Last Close Date:'+ rlc_latest_date
rlc_last_nav='NAV on Close: '+ "".join(map(str,rlc_last_nav))
rlc_change="Change in NAV:" + "["+str(rlcdatediff)+"]"

elss_latest_date='Last Close Date:'+ elss_latest_date
elss_last_nav='NAV on Close: '+ "".join(map(str,elss_last_nav))
elss_change="Change in NAV:" + "["+str(elssdatediff)+"]"

colors={
'background':'#111111',
'text':'#7FDBFF'
}


def change_style(change):
  if(change>0):
    style={'textAlign': 'left', 'color': 'green', 'size':18, 'font-weight': 'bold'}
  else:
   style={'textAlign': 'left', 'color': 'red', 'size':18, 'font-weight': 'bold'}
  return style

app.layout=html.Div(style={'backgroundColor': colors['background']},children=[

#Dash Image
#html.Img(src='data:image/png;base64,{}'.format(encoded_image.decode()),
#         style={'height': '7%','width': '7%', 'float' : 'right', 'position' : 'relative', 'padding-top' : 0, 'padding-right' : 0}),

#Dash Header
html.H3('Mutual Fund Portfolio',
style={'textAlign': 'center', 'color': colors['text'], 'font-weight': 'bold'}
      ),

#Body
html.Div([
 html.Div([
 html.Div("      "+ kotak_latest_date, style={ 'color': 'white', 'font-weight': 'bold'}),
 html.Div([html.Div("      "+ kotak_last_nav , style={ 'color': 'white', 'font-weight': 'bold'}),
 html.Div("      "+ kotak_change , style=change_style(kotakdatediff))] ),
 #, style={'width': '48%','float': 'left' ,'display': 'inline'}

 dcc.Graph(id='KotakStandard',
          figure={
             'data':[{'x':KotakStandard['Date'], 'y':KotakStandard['NAV'], 'type':'line', 'name':'KotakStandard'}],
             'layout':{'title': 'Kotak Standard Multicap Fund - Direct Plan - Growth',
                       'subtitle':kotak_latest_date,
                       'titlefont': {'size': 18, 'color':colors['text']},
                       #'xaxis':{'title': 'DATE', 'titlefont': {'size': 18, 'color':colors['text']}},
                       'yaxis':{'title': 'Net Asset Value', 'titlefont': {'size': 18, 'color':colors['text']}},
                       'plot_bgcolor':colors['background'],
                       'paper_bgcolor':colors['background'],
                       'font': {'color': '#F7BD81'} 
                       }
 
            }          
         )
        ],className="six columns")
 ,
 html.Div([
 html.Div("      "+ sbi_latest_date, style={'textAlign': 'left', 'color': 'white', 'font-weight': 'bold'}),
 html.Div([html.Div("      "+ sbi_last_nav , style={'textAlign': 'left', 'color': 'white', 'font-weight': 'bold'}),
           html.Div("      "+ sbi_change , style=change_style(sbidatediff))]),
 dcc.Graph(id='SBIBlueChip',
          figure={
             'data':[{'x':SBIBlueChip['Date'], 'y':SBIBlueChip['NAV'], 'type':'line', 'name':'SBIBlueChip'}],
             'layout':{'title':'SBI Blue Chip Fund - Growth',
                       'titlefont': {'size': 18, 'color':colors['text']},
                       #'xaxis':{'title': 'DATE', 'titlefont': {'size': 18, 'color':colors['text']}},
                       'yaxis':{'title': 'Net Asset Value', 'titlefont': {'size': 18, 'color':colors['text']}},
                       'plot_bgcolor':colors['background'],
                       'paper_bgcolor':colors['background'],
                       'font': {'color': '#F7BD81'}
                      }
                  }   
         )     
        ],className="six columns")

 ],className="row"),
html.Div([
 html.Div([
 html.Div("      "+ rlc_latest_date, style={'textAlign': 'left', 'color': 'white', 'font-weight': 'bold'}),
 html.Div([html.Div("      "+ rlc_last_nav , style={'textAlign': 'left', 'color': 'white', 'font-weight': 'bold'}),
           html.Div("      "+ rlc_change , style=change_style(rlcdatediff))]),
 dcc.Graph(id='RelainceSmallCap',
          figure={
             'data':[{'x':RelainceSmallCap['Date'], 'y':RelainceSmallCap['NAV'], 'type':'line', 'name':'RelainceSmallCap'}],
             'layout':{'title':'Reliance Small Cap Fund - Growth',
                       'titlefont': {'size': 18, 'color':colors['text']},
                       #'xaxis':{'title': 'DATE', 'titlefont': {'size': 18, 'color':colors['text']}},
                       'yaxis':{'title': 'Net Asset Value', 'titlefont': {'size': 18, 'color':colors['text']}},
                       'plot_bgcolor':colors['background'],
                       'paper_bgcolor':colors['background'],
                       'font': {'color': '#F7BD81'}
                       }
 
            }          
         )
        ],className="six columns")
 ,
 html.Div([
 html.Div("      "+ brl_latest_date, style={'textAlign': 'left', 'color': 'white', 'font-weight': 'bold'}),
 html.Div([html.Div("      "+ brl_last_nav , style={'textAlign': 'left', 'color': 'white', 'font-weight': 'bold'}),
           html.Div("      "+ brl_change , style=change_style(brldatediff))]),
 dcc.Graph(id='BirlaSmallCap',
          figure={
             'data':[{'x':BirlaSmallCap['Date'], 'y':BirlaSmallCap['NAV'], 'type':'line', 'name':'BirlaSmallCap'}],
             'layout':{'title':'Aditya Birla Sun Life Small cap Fund - Direct Plan - Growth',
                       'titlefont': {'size': 18, 'color':colors['text']},
                       #'xaxis':{'title': 'DATE', 'titlefont': {'size': 18, 'color':colors['text']}},
                       'yaxis':{'title': 'Net Asset Value', 'titlefont': {'size': 18, 'color':colors['text']}},
                       'plot_bgcolor':colors['background'],
                       'paper_bgcolor':colors['background'],
                       'font': {'color': '#F7BD81'}  
                       }
                  }   
         )     
        ],className="six columns")

 ],className="row")
,

html.Div([
 html.Div("      "+ elss_latest_date, style={'textAlign': 'left', 'color': 'white', 'font-weight': 'bold'}),
 html.Div([html.Div("      "+ elss_last_nav , style={'textAlign': 'left', 'color': 'white', 'font-weight': 'bold'}),
           html.Div("      "+ elss_change , style=change_style(elssdatediff))]),
 dcc.Graph(id='BirlaELSS',
          figure={
             'data':[{'x':BirlaELSS['Date'], 'y':BirlaELSS['NAV'], 'type':'line', 'name':'BirlaELSS'}],
             'layout':{'title':'Aditya Birla Sun Life Tax Relief 96 - Regular Plan - Growth',
                       'titlefont': {'size': 18, 'color':colors['text']},
                       #'xaxis':{'title': 'DATE', 'titlefont': {'size': 18, 'color':colors['text']}},
                       'yaxis':{'title': 'Net Asset Value', 'titlefont': {'size': 18, 'color':colors['text']}},
                       'plot_bgcolor':colors['background'],
                       'paper_bgcolor':colors['background'],
                       'font': {'color': '#F7BD81'} 
                       }
 
            }          
         )
        ])

 ])
app.css.append_css({
    'external_url': 'https://codepen.io/chriddyp/pen/bWLwgP.css'
})

if __name__ =='__main__':
 app.run_server(debug=True)