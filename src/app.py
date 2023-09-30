import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
import win32print
import win32ui
from PIL import Image, ImageWin
import os
import datetime
from barcode import generate
from barcode.writer import ImageWriter

app = dash.Dash(__name__)
server = app.server
try:
    printer_name = win32print.GetDefaultPrinter()
except:
    printer_name= "There is no printed detected :("
# Define the layout of the app
app.layout = html.Div(
    [
    html.Div(
        html.H1("Click the button to print 8 barcodes"),
        style={
            'backgroundColor': 'black',
            'color': 'white',
            'padding': '20px',
            'textAlign': 'center',
        }
    ),

    html.Div([
        dcc.Input(
            id='text-input',
            type='text', 
            placeholder='Enter you barcode data here, to be printed in: '+ printer_name,
            style={
                'textAlign': 'center',
                # 'margin-right': '2%',
                'width':"100%",
                'display':"flex"
            }
        ),
        html.Br(),
        html.Button(
            "Print Barcodes",
            id="print-button",
            n_clicks=0,
            style={
                'textAlign': 'center',
                'margin-left': '2%',
                
                'display':"flex"
            }
        ),
        html.Br(),
        html.Br(),
        html.Br(),

        html.Div(id="output-message", children="")
    ],style={'textAlign': 'center', 'margin': '20px'})
 
])

# Define a callback to handle button click
@app.callback(
    Output("output-message", "children"),
    Input("print-button", "n_clicks"),
    State("text-input", "value")
)
def print_barcodes(n_clicks, barcode_string):
    if n_clicks > 0:
        PHYSICALWIDTH = 110
        PHYSICALHEIGHT = 111
        
        current_date = datetime.datetime.now().strftime("%Y_%m_%d__%H_%M_%S")
        barcode_image_output = os.path.join('barcodes\images', 'barcode_image'+ current_date)
        barcode_image = generate(
                                    'code128', 
                                    str(barcode_string), 
                                    writer=ImageWriter(), 
                                    writer_options={"text_distance": 1.2, "font_size": 20, "module_width":0.3 }, 
                                    output=barcode_image_output
                                )
        
        printer_name = win32print.GetDefaultPrinter()
        print(printer_name)
        file_name = r"C:\Users\fella\Desktop\Dready\CEMIX\src\barcodes\images\barcode_image"+ current_date+".png"

        for i in range(8):
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)
            printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)

            bmp = Image.open (file_name)
            if bmp.size[0] < bmp.size[1]:
                bmp = bmp.rotate (90)

            hDC.StartDoc (file_name)
            hDC.StartPage ()

            dib = ImageWin.Dib (bmp)
            x, y = 0, 0
            dib.draw (hDC.GetHandleOutput(), (x,y,printer_size[0],printer_size[1]))

            hDC.EndPage ()
            hDC.EndDoc ()
            hDC.DeleteDC ()
        

        barcode_message = "Barcodes printed successfully!"
        return barcode_message

if __name__ == '__main__':
    app.run_server(debug=True)
